VERSION 2.00
Begin Form frmSetup 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "SPG-NEWS環境設定"
   ClientHeight    =   9015
   ClientLeft      =   885
   ClientTop       =   1485
   ClientWidth     =   13545
   Height          =   9420
   Left            =   825
   LinkTopic       =   "Form1"
   ScaleHeight     =   9015
   ScaleWidth      =   13545
   Top             =   1140
   Width           =   13665
   Begin Tack tkSetup 
      BackColor       =   &H00C0C0C0&
      Caption         =   "個人情報;通信;動作ﾓｰﾄﾞ;ﾃﾞｨﾚｸﾄﾘ;外部ﾌﾟﾛｸﾞﾗﾑ"
      Escapement      =   0  '文字回転なし
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "FixedSys"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   8835
      Left            =   0
      TabIndex        =   0
      Tack3D          =   -1  'True
      TackAlign       =   0  '上付タック
      TackIndex       =   0
      TackMax         =   5
      TackSize        =   8
      Top             =   0
      Width           =   13335
      Begin SSFrame fraSetup 
         Caption         =   "個人情報"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   4995
         Index           =   0
         Left            =   10800
         TabIndex        =   37
         Top             =   7080
         Width           =   8535
         Begin SSFrame Frame3D1 
            Caption         =   "TimeZone"
            FontBold        =   0   'False
            FontItalic      =   0   'False
            FontName        =   "FixedSys"
            FontSize        =   13.5
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   975
            Left            =   360
            TabIndex        =   46
            Top             =   2700
            Width           =   2955
            Begin ComboBox cboTimeZone 
               FontBold        =   0   'False
               FontItalic      =   0   'False
               FontName        =   "FixedSys"
               FontSize        =   13.5
               FontStrikethru  =   0   'False
               FontUnderline   =   0   'False
               Height          =   390
               Left            =   360
               Style           =   2  'Dropdown List
               TabIndex        =   47
               Top             =   420
               Width           =   2295
            End
         End
         Begin CommandButton cmdOK 
            Caption         =   "OK"
            FontBold        =   0   'False
            FontItalic      =   0   'False
            FontName        =   "FixedSys"
            FontSize        =   13.5
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            Height          =   495
            Index           =   0
            Left            =   3900
            TabIndex        =   45
            Top             =   4140
            Width           =   1215
         End
         Begin CommandButton cmdCancel 
            Caption         =   "ｷｬﾝｾﾙ"
            FontBold        =   0   'False
            FontItalic      =   0   'False
            FontName        =   "FixedSys"
            FontSize        =   13.5
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            Height          =   495
            Index           =   0
            Left            =   5340
            TabIndex        =   44
            Top             =   4140
            Width           =   1215
         End
         Begin SSFrame fraName 
            Caption         =   "使用者名"
            FontBold        =   0   'False
            FontItalic      =   0   'False
            FontName        =   "FixedSys"
            FontSize        =   13.5
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   975
            Left            =   360
            TabIndex        =   42
            Top             =   1560
            Width           =   7155
            Begin TextBox txtMyName 
               FontBold        =   0   'False
               FontItalic      =   0   'False
               FontName        =   "FixedSys"
               FontSize        =   13.5
               FontStrikethru  =   0   'False
               FontUnderline   =   0   'False
               Height          =   495
               Left            =   240
               TabIndex        =   43
               Top             =   360
               Width           =   6795
            End
         End
         Begin SSFrame fraHomeRBBS 
            Caption         =   "ホームＲＢＢＳ"
            FontBold        =   0   'False
            FontItalic      =   0   'False
            FontName        =   "FixedSys"
            FontSize        =   13.5
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   915
            Left            =   3360
            TabIndex        =   40
            Top             =   540
            Width           =   4155
            Begin TextBox txtFwdCall 
               FontBold        =   0   'False
               FontItalic      =   0   'False
               FontName        =   "FixedSys"
               FontSize        =   13.5
               FontStrikethru  =   0   'False
               FontUnderline   =   0   'False
               Height          =   495
               Left            =   120
               TabIndex        =   41
               Text            =   "FWDCall"
               Top             =   300
               Width           =   3915
            End
         End
         Begin SSFrame fraMyCall 
            Caption         =   "コールサイン"
            FontBold        =   0   'False
            FontItalic      =   0   'False
            FontName        =   "FixedSys"
            FontSize        =   13.5
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   915
            Left            =   360
            TabIndex        =   38
            Top             =   540
            Width           =   2895
            Begin TextBox txtMyCall 
               BackColor       =   &H00FFFFFF&
               FontBold        =   0   'False
               FontItalic      =   0   'False
               FontName        =   "FixedSys"
               FontSize        =   13.5
               FontStrikethru  =   0   'False
               FontUnderline   =   0   'False
               Height          =   495
               Left            =   120
               TabIndex        =   39
               Text            =   "MyCall"
               Top             =   300
               Width           =   2655
            End
         End
      End
      Begin SSFrame fraSetup 
         Caption         =   "外部プログラム設定"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   3795
         Index           =   4
         Left            =   660
         TabIndex        =   27
         Top             =   3480
         Visible         =   0   'False
         Width           =   7935
         Begin CommonDialog CMDlgEditor 
            DefaultExt      =   "TXT"
            Left            =   7140
            Top             =   420
         End
         Begin SSFrame fraEditor 
            Caption         =   "テキストエディタ"
            FontBold        =   0   'False
            FontItalic      =   0   'False
            FontName        =   "FixedSys"
            FontSize        =   13.5
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   1095
            Left            =   660
            TabIndex        =   35
            Top             =   660
            Width           =   5775
            Begin TextBox txtEditor 
               FontBold        =   0   'False
               FontItalic      =   0   'False
               FontName        =   "FixedSys"
               FontSize        =   13.5
               FontStrikethru  =   0   'False
               FontUnderline   =   0   'False
               Height          =   495
               Left            =   120
               TabIndex        =   24
               TabStop         =   0   'False
               Top             =   420
               Width           =   3195
            End
            Begin CommandButton cmdEditor 
               Caption         =   "参照"
               FontBold        =   0   'False
               FontItalic      =   0   'False
               FontName        =   "FixedSys"
               FontSize        =   13.5
               FontStrikethru  =   0   'False
               FontUnderline   =   0   'False
               Height          =   375
               Left            =   4080
               TabIndex        =   36
               TabStop         =   0   'False
               Top             =   480
               Width           =   1215
            End
         End
         Begin CommandButton cmdOK 
            Caption         =   "OK"
            Default         =   -1  'True
            FontBold        =   0   'False
            FontItalic      =   0   'False
            FontName        =   "FixedSys"
            FontSize        =   13.5
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            Height          =   495
            Index           =   4
            Left            =   5280
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   3000
            Width           =   1215
         End
         Begin CommandButton cmdCancel 
            Caption         =   "ｷｬﾝｾﾙ"
            FontBold        =   0   'False
            FontItalic      =   0   'False
            FontName        =   "FixedSys"
            FontSize        =   13.5
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            Height          =   495
            Index           =   4
            Left            =   6660
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   3000
            Width           =   1215
         End
      End
      Begin SSFrame fraSetup 
         Caption         =   "ディレクトリ"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   3555
         Index           =   3
         Left            =   3840
         TabIndex        =   3
         Top             =   5400
         Visible         =   0   'False
         Width           =   8175
         Begin SSFrame fraMsgDir 
            Caption         =   "メッセージ格納ディレクトリ"
            FontBold        =   0   'False
            FontItalic      =   0   'False
            FontName        =   "FixedSys"
            FontSize        =   13.5
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   2775
            Left            =   360
            TabIndex        =   34
            Top             =   360
            Width           =   5115
            Begin DirListBox dirMsg 
               FontBold        =   0   'False
               FontItalic      =   0   'False
               FontName        =   "FixedSys"
               FontSize        =   13.5
               FontStrikethru  =   0   'False
               FontUnderline   =   0   'False
               Height          =   1230
               Left            =   300
               TabIndex        =   21
               TabStop         =   0   'False
               Top             =   1380
               Width           =   3075
            End
            Begin DriveListBox drvMsg 
               FontBold        =   0   'False
               FontItalic      =   0   'False
               FontName        =   "FixedSys"
               FontSize        =   13.5
               FontStrikethru  =   0   'False
               FontUnderline   =   0   'False
               Height          =   420
               Left            =   300
               TabIndex        =   20
               TabStop         =   0   'False
               Top             =   900
               Width           =   3075
            End
            Begin TextBox txtMsgDir 
               FontBold        =   0   'False
               FontItalic      =   0   'False
               FontName        =   "FixedSys"
               FontSize        =   13.5
               FontStrikethru  =   0   'False
               FontUnderline   =   0   'False
               Height          =   495
               Left            =   300
               TabIndex        =   19
               TabStop         =   0   'False
               Text            =   "C:\MSG"
               Top             =   360
               Width           =   4575
            End
         End
         Begin CommandButton cmdOK 
            Caption         =   "OK"
            FontBold        =   0   'False
            FontItalic      =   0   'False
            FontName        =   "FixedSys"
            FontSize        =   13.5
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            Height          =   495
            Index           =   3
            Left            =   5760
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   1380
            Width           =   1215
         End
         Begin CommandButton cmdCancel 
            Caption         =   "ｷｬﾝｾﾙ"
            FontBold        =   0   'False
            FontItalic      =   0   'False
            FontName        =   "FixedSys"
            FontSize        =   13.5
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            Height          =   495
            Index           =   3
            Left            =   6300
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   2520
            Width           =   1215
         End
      End
      Begin SSFrame fraSetup 
         Caption         =   "動作モード"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   2895
         Index           =   2
         Left            =   360
         TabIndex        =   2
         Top             =   5640
         Visible         =   0   'False
         Width           =   7035
         Begin SSOption optMode 
            Caption         =   "ＡＴＥＲＭ互換"
            Enabled         =   0   'False
            FontBold        =   0   'False
            FontItalic      =   0   'False
            FontName        =   "FixedSys"
            FontSize        =   13.5
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            Height          =   495
            Index           =   3
            Left            =   420
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   1920
            Visible         =   0   'False
            Width           =   2295
         End
         Begin SSOption optMode 
            Caption         =   "ＤＪＬ－ＭＢ互換"
            FontBold        =   0   'False
            FontItalic      =   0   'False
            FontName        =   "FixedSys"
            FontSize        =   13.5
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            Height          =   495
            Index           =   2
            Left            =   420
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   1440
            Width           =   2295
         End
         Begin SSOption optMode 
            Caption         =   "オートパイロット"
            FontBold        =   0   'False
            FontItalic      =   0   'False
            FontName        =   "FixedSys"
            FontSize        =   13.5
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            Height          =   495
            Index           =   1
            Left            =   420
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   960
            Width           =   2295
         End
         Begin SSOption optMode 
            Caption         =   "ＦＷＤ転送"
            FontBold        =   0   'False
            FontItalic      =   0   'False
            FontName        =   "FixedSys"
            FontSize        =   13.5
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            Height          =   495
            Index           =   0
            Left            =   420
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   480
            Width           =   1695
         End
         Begin CommandButton cmdOK 
            Caption         =   "OK"
            FontBold        =   0   'False
            FontItalic      =   0   'False
            FontName        =   "FixedSys"
            FontSize        =   13.5
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            Height          =   495
            Index           =   2
            Left            =   4200
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   2280
            Width           =   1215
         End
         Begin CommandButton cmdCancel 
            Caption         =   "ｷｬﾝｾﾙ"
            FontBold        =   0   'False
            FontItalic      =   0   'False
            FontName        =   "FixedSys"
            FontSize        =   13.5
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            Height          =   495
            Index           =   2
            Left            =   5580
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   2280
            Width           =   1215
         End
      End
      Begin SSFrame fraSetup 
         Caption         =   "通信設定"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "FixedSys"
         FontSize        =   13.5
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   4815
         Index           =   1
         Left            =   1320
         TabIndex        =   1
         Top             =   960
         Visible         =   0   'False
         Width           =   9735
         Begin SSFrame Frame3D2 
            Caption         =   "フロー制御"
            FontBold        =   0   'False
            FontItalic      =   0   'False
            FontName        =   "System"
            FontSize        =   13.5
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   735
            Left            =   3360
            TabIndex        =   52
            Top             =   2400
            Width           =   3255
            Begin SSOption optFlow 
               Caption         =   "ｿﾌﾄﾌﾛｰ"
               FontBold        =   0   'False
               FontItalic      =   0   'False
               FontName        =   "System"
               FontSize        =   13.5
               FontStrikethru  =   0   'False
               FontUnderline   =   0   'False
               Height          =   255
               Index           =   1
               Left            =   1560
               TabIndex        =   54
               Top             =   360
               Width           =   1215
            End
            Begin SSOption optFlow 
               Caption         =   "ﾊｰﾄﾞﾌﾛｰ"
               FontBold        =   0   'False
               FontItalic      =   0   'False
               FontName        =   "System"
               FontSize        =   13.5
               FontStrikethru  =   0   'False
               FontUnderline   =   0   'False
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   53
               Top             =   360
               Width           =   1215
            End
         End
         Begin SSFrame fraDelay 
            Caption         =   "送信ディレイ"
            FontBold        =   0   'False
            FontItalic      =   0   'False
            FontName        =   "FixedSys"
            FontSize        =   13.5
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   975
            Left            =   240
            TabIndex        =   50
            Top             =   2640
            Width           =   2955
            Begin TextBox txtSendDelay 
               FontBold        =   0   'False
               FontItalic      =   0   'False
               FontName        =   "FixedSys"
               FontSize        =   13.5
               FontStrikethru  =   0   'False
               FontUnderline   =   0   'False
               Height          =   495
               Left            =   240
               TabIndex        =   51
               Text            =   "Text1"
               Top             =   360
               Width           =   2415
            End
         End
         Begin SSFrame fraTimeOut 
            Caption         =   "通信ﾀｲﾑｱｳﾄ時間"
            FontBold        =   0   'False
            FontItalic      =   0   'False
            FontName        =   "FixedSys"
            FontSize        =   13.5
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   1995
            Left            =   5220
            TabIndex        =   30
            Top             =   360
            Width           =   2295
            Begin ComboBox cboSendTimeOut 
               FontBold        =   0   'False
               FontItalic      =   0   'False
               FontName        =   "FixedSys"
               FontSize        =   13.5
               FontStrikethru  =   0   'False
               FontUnderline   =   0   'False
               Height          =   390
               Left            =   240
               Style           =   2  'Dropdown List
               TabIndex        =   12
               TabStop         =   0   'False
               Top             =   1440
               Width           =   1455
            End
            Begin ComboBox cboRcvTimeOut 
               FontBold        =   0   'False
               FontItalic      =   0   'False
               FontName        =   "FixedSys"
               FontSize        =   13.5
               FontStrikethru  =   0   'False
               FontUnderline   =   0   'False
               Height          =   390
               Left            =   240
               Style           =   2  'Dropdown List
               TabIndex        =   11
               TabStop         =   0   'False
               Top             =   660
               Width           =   1455
            End
            Begin Label Label2 
               BackColor       =   &H00C0C0C0&
               Caption         =   "送信ﾀｲﾑｱｳﾄ(分)"
               FontBold        =   0   'False
               FontItalic      =   0   'False
               FontName        =   "FixedSys"
               FontSize        =   13.5
               FontStrikethru  =   0   'False
               FontUnderline   =   0   'False
               Height          =   315
               Left            =   240
               TabIndex        =   28
               Top             =   1140
               Width           =   1935
            End
            Begin Label Label1 
               BackColor       =   &H00C0C0C0&
               Caption         =   "受信ﾀｲﾑｱｳﾄ(分)"
               FontBold        =   0   'False
               FontItalic      =   0   'False
               FontName        =   "FixedSys"
               FontSize        =   13.5
               FontStrikethru  =   0   'False
               FontUnderline   =   0   'False
               Height          =   315
               Left            =   240
               TabIndex        =   29
               Top             =   360
               Width           =   1995
            End
         End
         Begin SSFrame fraPort 
            Caption         =   "通信ポート"
            FontBold        =   0   'False
            FontItalic      =   0   'False
            FontName        =   "FixedSys"
            FontSize        =   13.5
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   1935
            Left            =   3360
            TabIndex        =   32
            Top             =   360
            Width           =   1695
            Begin SSOption optPort 
               Caption         =   "COM4"
               Enabled         =   0   'False
               FontBold        =   0   'False
               FontItalic      =   0   'False
               FontName        =   "FixedSys"
               FontSize        =   13.5
               FontStrikethru  =   0   'False
               FontUnderline   =   0   'False
               Height          =   375
               Index           =   3
               Left            =   240
               TabIndex        =   10
               TabStop         =   0   'False
               Top             =   1440
               Width           =   1215
            End
            Begin SSOption optPort 
               Caption         =   "COM3"
               Enabled         =   0   'False
               FontBold        =   0   'False
               FontItalic      =   0   'False
               FontName        =   "FixedSys"
               FontSize        =   13.5
               FontStrikethru  =   0   'False
               FontUnderline   =   0   'False
               Height          =   375
               Index           =   2
               Left            =   240
               TabIndex        =   9
               TabStop         =   0   'False
               Top             =   1080
               Width           =   1215
            End
            Begin SSOption optPort 
               Caption         =   "COM2"
               Enabled         =   0   'False
               FontBold        =   0   'False
               FontItalic      =   0   'False
               FontName        =   "FixedSys"
               FontSize        =   13.5
               FontStrikethru  =   0   'False
               FontUnderline   =   0   'False
               Height          =   375
               Index           =   1
               Left            =   240
               TabIndex        =   8
               TabStop         =   0   'False
               Top             =   720
               Width           =   1215
            End
            Begin SSOption optPort 
               Caption         =   "COM1"
               Enabled         =   0   'False
               FontBold        =   0   'False
               FontItalic      =   0   'False
               FontName        =   "FixedSys"
               FontSize        =   13.5
               FontStrikethru  =   0   'False
               FontUnderline   =   0   'False
               Height          =   375
               Index           =   0
               Left            =   240
               TabIndex        =   7
               TabStop         =   0   'False
               Top             =   360
               Width           =   1215
            End
         End
         Begin SSFrame fraBaud 
            Caption         =   "通信速度"
            FontBold        =   0   'False
            FontItalic      =   0   'False
            FontName        =   "FixedSys"
            FontSize        =   13.5
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   975
            Left            =   240
            TabIndex        =   31
            Top             =   1560
            Width           =   2955
            Begin ComboBox cboBaud 
               FontBold        =   0   'False
               FontItalic      =   0   'False
               FontName        =   "FixedSys"
               FontSize        =   13.5
               FontStrikethru  =   0   'False
               FontUnderline   =   0   'False
               Height          =   390
               Left            =   240
               Style           =   2  'Dropdown List
               TabIndex        =   6
               TabStop         =   0   'False
               Top             =   300
               Width           =   2415
            End
         End
         Begin SSFrame fraLink 
            Caption         =   "接続方法"
            FontBold        =   0   'False
            FontItalic      =   0   'False
            FontName        =   "FixedSys"
            FontSize        =   13.5
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   1035
            Left            =   240
            TabIndex        =   33
            Top             =   360
            Width           =   2955
            Begin SSOption optLink 
               Caption         =   "直結"
               Enabled         =   0   'False
               FontBold        =   0   'False
               FontItalic      =   0   'False
               FontName        =   "FixedSys"
               FontSize        =   13.5
               FontStrikethru  =   0   'False
               FontUnderline   =   0   'False
               Height          =   555
               Index           =   1
               Left            =   1200
               TabIndex        =   5
               TabStop         =   0   'False
               Top             =   300
               Width           =   1695
            End
            Begin SSOption optLink 
               Caption         =   "TNC"
               Enabled         =   0   'False
               FontBold        =   0   'False
               FontItalic      =   0   'False
               FontName        =   "FixedSys"
               FontSize        =   13.5
               FontStrikethru  =   0   'False
               FontUnderline   =   0   'False
               Height          =   555
               Index           =   0
               Left            =   240
               TabIndex        =   4
               TabStop         =   0   'False
               Top             =   300
               Width           =   1215
            End
         End
         Begin CommandButton cmdOK 
            Caption         =   "OK"
            FontBold        =   0   'False
            FontItalic      =   0   'False
            FontName        =   "FixedSys"
            FontSize        =   13.5
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            Height          =   495
            Index           =   1
            Left            =   6840
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   2880
            Width           =   1215
         End
         Begin CommandButton cmdCancel 
            Caption         =   "ｷｬﾝｾﾙ"
            FontBold        =   0   'False
            FontItalic      =   0   'False
            FontName        =   "FixedSys"
            FontSize        =   13.5
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            Height          =   495
            Index           =   1
            Left            =   8160
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   2880
            Width           =   1215
         End
      End
   End
End
Option Explicit

Sub cmdCancel_Click (Index As Integer)
    If FIRST = True Then
	MsgBox "初期設定はキャンセルされました。", MB_ICONINFORMATION
	End
    End If
    Unload frmSetup
End Sub

Sub cmdEditor_Click ()
    Dim s$

    s$ = frmSetup!txtEditor
    On Error GoTo ERRORPROC
    CMDlgEditor.Flags = &H4
    CMDlgEditor.DefaultExt = "EXE"
    CMDlgEditor.DialogTitle = "外部エディタの設定"
    CMDlgEditor.Filter = "EXE File (*.EXE)|*.EXE|"
    ' 最初に使うフィルタを“All Files”に設定します.
    CMDlgEditor.FilterIndex = 1
    ' [上書き禁止] チェック ボックスをチェックします.
'    CMDlgEditor.Flags = OFN_READONLY
    CMDlgEditor.CancelError = -1
    ' [ファイルを開く] ダイアログ ボックスを表示します.
    CMDlgEditor.Action = 1
    ' ユーザーが <OK> ボタンを選択したときの処理
    frmSetup!txtEditor = CMDlgEditor.Filename
    ChDrive Left$(SYSDIR, 2)
    ChDir Left$(SYSDIR, Len(SYSDIR) - 1)
    Exit Sub

ERRORPROC:
    frmSetup!txtEditor = s$
    ChDrive Left$(SYSDIR, 2)
    ChDir Left$(SYSDIR, Len(SYSDIR) - 1)
    Exit Sub
End Sub

Sub cmdOK_Click (Index As Integer)
    Dim i, a As Integer
    Dim p$, rt$, st$, msg$
    Dim Filenum As Integer
    Dim Contents As String

    ' CONFIGダイアログ設定
    MYCALL = UCase$(txtMyCall.Text)
    HOMERBBS = UCase$(txtFwdCall.Text)
    MYNAME = txtMyName.Text
    TIMEZONE = cboTimeZone.Text
    If frmSetup!optLink(0).Value = True Then
	LINK = "TNC"
    Else
	LINK = "RS232C"
    End If
    BAUD = frmSetup!cboBaud.Text
    msg$ = UCase$(frmSetup!txtMsgDir.Text)
    EDITOR = UCase$(frmSetup!txtEditor.Text)
    For i = 0 To 4
	If frmSetup!optPort(i).Value = True Then
	    Port = i + 1
	    Exit For
	End If
    Next i
    If frmSetup!optMode(0).Value = True Then
	MODE = "FWD"
    ElseIf frmSetup!optMode(1).Value = True Then
	MODE = "AUTO"
    ElseIf frmSetup!optMode(2).Value = True Then
	MODE = "DJLMB"
    ElseIf frmSetup!optMode(3).Value = True Then
	MODE = "ATERM"
    Else
	MODE = "FWD"
    End If
    rt$ = cboRcvTimeOut.Text
    st$ = cboSendTimeOut.Text
    RCVWAITTIME = Val(rt$)
    SENDWAITTIME = Val(st$)
    SENDDELAY = Val(txtSendDelay.Text)
    If frmSetup!optFlow(0).Value = True Then
	XFLOWFLAG = 0
    Else
	XFLOWFLAG = 1
    End If

' 環境設定ファイル更新
    p$ = Format$(Port)
    a = WritePrivateProfileString("Config", "MyCall", MYCALL + Chr$(0), INIFILE)
    a = WritePrivateProfileString("Config", "HomeRBBS", HOMERBBS + Chr$(0), INIFILE)
    a = WritePrivateProfileString("Config", "MyName", MYNAME + Chr$(0), INIFILE)
    a = WritePrivateProfileString("Config", "Link", LINK + Chr$(0), INIFILE)
    a = WritePrivateProfileString("Config", "Baud", BAUD + Chr$(0), INIFILE)
    a = WritePrivateProfileString("Config", "MsgDir", msg$ + Chr$(0), INIFILE)
    a = WritePrivateProfileString("Config", "Editor", EDITOR + Chr$(0), INIFILE)
    a = WritePrivateProfileString("Config", "Port", p$ + Chr$(0), INIFILE)
    a = WritePrivateProfileString("Config", "Mode", MODE + Chr$(0), INIFILE)
    a = WritePrivateProfileString("Config", "RcvTimeOut", rt$ + Chr$(0), INIFILE)
    a = WritePrivateProfileString("Config", "SendTimeOut", st$ + Chr$(0), INIFILE)
    a = WritePrivateProfileString("Config", "SendDelay", txtSendDelay.Text + Chr$(0), INIFILE)
    a = WritePrivateProfileString("Config", "TimeZone", TIMEZONE + Chr$(0), INIFILE)
    a = WritePrivateProfileString("Config", "XFlow", Format$(XFLOWFLAG), INIFILE)
    
    If FIRST = True Then
	MsgBox "ＳＰＧ－ＮＥＷＳを再起動して下さい。", MB_ICONINFORMATION
	End
    End If

    If L_MODE = 0 Then
	MsgBox "メッセージを格納するディレクトリまたは言語が変更された場合、" + Chr$(13) + Chr$(10) + "その変更は次回起動時より有効となります。", MB_ICONINFORMATION
    Else
	MsgBox "When reactivating ,the change in the directory or the language becomes effective."
    End If
    
    If MODE = "FWD" Then
	frmMain!mnuComSR.Enabled = True
	frmMain!cmdFWD.Enabled = True
	frmMain!cmdAuto.Enabled = False
	frmMain!mnuAutoPilot.Enabled = False
	frmMain!mnuAutoPilotConfig.Enabled = False
    ElseIf MODE = "AUTO" Then
	frmMain!mnuComSR.Enabled = False
	frmMain!cmdFWD.Enabled = False
	frmMain!cmdAuto.Enabled = True
	frmMain!mnuComSR.Enabled = True
	frmMain!mnuAutoPilot.Enabled = True
	frmMain!mnuAutoPilotConfig.Enabled = True
    ElseIf MODE = "DJLMB" Then
	frmMain!mnuComSR.Enabled = False
	frmMain!cmdFWD.Enabled = False
	frmMain!cmdAuto.Enabled = False
	frmMain!mnuComSR.Enabled = False
	frmMain!mnuAutoPilot.Enabled = False
	frmMain!mnuAutoPilotConfig.Enabled = False
    End If

    Unload frmSetup
End Sub

Sub dirMsg_Change ()
    txtMsgDir.Text = dirMsg.Path
End Sub

Sub drvMsg_Change ()
    On Error GoTo err1
    dirMsg.Path = drvMsg.Drive
err1:
    Resume Next
End Sub

Sub EnglishSet ()
    Dim i As Integer

    frmSetup.Caption = "SPG-NEWS Setup"
    tkSetup.Caption = "Personal;Communication;Mode;Directory;Ext Program"
    '
    fraSetup(0).Caption = "Personal"
    fraSetup(1).Caption = "Communication"
    fraSetup(2).Caption = "Mode"
    fraSetup(3).Caption = "Directory"
    fraSetup(4).Caption = "External Program"
    '
    fraMyCall.Caption = "Callsign"
    fraHomeRBBS.Caption = "Home RBBS"
    fraName.Caption = "Name"
    '
    For i = 0 To 4
	cmdCancel(i).Caption = "Cancel"
    Next i
    '
    fraLink.Caption = "Connect Method"
    fraBaud.Caption = "Speed"
    fraPort.Caption = "Com Port"
    fraTimeOut.Caption = "Time out(min)"
    optLink(0).Caption = "TNC"
    optLink(1).Caption = "Serial Link"
    Label1.Caption = "Rcv"
    Label2.Caption = "Send"
    '
    optMode(0).Caption = "Forward Protocol"
    optMode(1).Caption = "Auto Pilot"
    '
    fraMsgDIr.Caption = "Message Directory"
    '
    fraEditor.Caption = "External Editor"
    cmdEditor.Caption = "Reference"
End Sub

Sub Form_Load ()
    Dim Filenum As Integer
    Dim i%, s$, r$

    On Error Resume Next

    Width = 8000
    Height = 5000
    Top = (screen.Height - Height) / 2
    Left = (screen.Width - Width) / 2
    tkSetup.Top = 0
    tkSetup.Left = 0
    tkSetup.Height = frmSetup.ScaleHeight
    tkSetup.Width = frmSetup.ScaleWidth
    
    fraSetup(0).Top = 480
    fraSetup(0).Left = 100
    fraSetup(0).Height = tkSetup.Height - 480 - 100
    fraSetup(0).Width = tkSetup.Width - 100 - 100
    fraSetup(0).ZOrder
    cmdOK(0).Top = fraSetup(0).Height - (cmdOK(0).Height + 120)
    cmdCancel(0).Top = cmdOK(0).Top
    cmdOK(0).Left = fraSetup(0).Width - (cmdOK(0).Width + 120 + cmdCancel(0).Width + 480)
    cmdCancel(0).Left = cmdOK(0).Left + cmdOK(0).Width + 120
    cmdOK(0).Default = True
    fraSetup(0).Visible = True

    For i% = 0 To frmSetup.Controls.Count - 1
	frmSetup.Controls(i).TabStop = False
    Next i%
    
    tkSetup.TabIndex = 0
    txtMyCall.TabIndex = 1
    txtFwdCall.TabIndex = 2
    txtMyName.TabIndex = 3
    cboTimeZone.TabIndex = 4
    cmdOK(0).TabIndex = 5
    cmdCancel(0).TabIndex = 6

    tkSetup.TabStop = True
    txtMyCall.TabStop = True
    txtFwdCall.TabStop = True
    txtMyName.TabStop = True
    cboTimeZone.TabStop = True
    cmdOK(0).TabStop = True
    cmdCancel(0).TabStop = True
    
    frmSetup!cboBaud.AddItem "38400"
    frmSetup!cboBaud.AddItem "19200"
    frmSetup!cboBaud.AddItem "9600"
    frmSetup!cboBaud.AddItem "4800"
    frmSetup!cboBaud.AddItem "2400"
    frmSetup!cboBaud.AddItem "1200"

    frmSetup!cboTimeZone.AddItem "+0000"
    frmSetup!cboTimeZone.AddItem "-0100"
    frmSetup!cboTimeZone.AddItem "-0200"
    frmSetup!cboTimeZone.AddItem "-0300"
    frmSetup!cboTimeZone.AddItem "-0400"
    frmSetup!cboTimeZone.AddItem "-0500"
    frmSetup!cboTimeZone.AddItem "-0600"
    frmSetup!cboTimeZone.AddItem "-0700"
    frmSetup!cboTimeZone.AddItem "-0800"
    frmSetup!cboTimeZone.AddItem "-0900"
    frmSetup!cboTimeZone.AddItem "-1000"
    frmSetup!cboTimeZone.AddItem "-1100"
    frmSetup!cboTimeZone.AddItem "-1200"
    frmSetup!cboTimeZone.AddItem "+0100"
    frmSetup!cboTimeZone.AddItem "+0200"
    frmSetup!cboTimeZone.AddItem "+0300"
    frmSetup!cboTimeZone.AddItem "+0400"
    frmSetup!cboTimeZone.AddItem "+0500"
    frmSetup!cboTimeZone.AddItem "+0600"
    frmSetup!cboTimeZone.AddItem "+0700"
    frmSetup!cboTimeZone.AddItem "+0800"
    frmSetup!cboTimeZone.AddItem "+0900"
    frmSetup!cboTimeZone.AddItem "+1000"
    frmSetup!cboTimeZone.AddItem "+1100"
    frmSetup!cboTimeZone.AddItem "+1200"
    
    If XFLOWFLAG = 0 Then
	frmSetup!optFlow(0).Value = True
    Else
	frmSetup!optFlow(1).Value = True
    End If

    For i% = 1 To 30
	frmSetup!cboRcvTimeOut.AddItem Format$(i%)
	frmSetup!cboSendTimeOut.AddItem Format$(i%)
    Next i%
    
    If L_MODE = 1 Then
	EnglishSet
    End If
End Sub

Sub tkSetup_Change ()
    Dim tabnum%, i%
    
    On Error GoTo tkSetup_ERROR

    For i% = 0 To 4
	fraSetup(tabnum%).Enabled = False
    Next i%
    
    For i% = 0 To frmSetup.Controls.Count - 1
10      frmSetup.Controls(i).TabStop = False
    Next i%
    tkSetup.TabStop = True
	
    optLink(0).Enabled = False
    optLink(1).Enabled = False
    optPort(0).Enabled = False
    optPort(1).Enabled = False
    optPort(2).Enabled = False
    optPort(3).Enabled = False

    fraSetup(tabnum%).Enabled = True
    tabnum% = tkSetup.TackIndex
    fraSetup(tabnum%).Top = 480
    fraSetup(tabnum%).Left = 100
    fraSetup(tabnum%).Height = tkSetup.Height - 480 - 100
    fraSetup(tabnum%).Width = tkSetup.Width - 100 - 100
    fraSetup(tabnum%).ZOrder
    cmdOK(tabnum%).Top = fraSetup(tabnum%).Height - (cmdOK(tabnum%).Height + 120)
    cmdCancel(tabnum%).Top = cmdOK(tabnum%).Top
    cmdOK(tabnum%).Left = fraSetup(tabnum%).Width - (cmdOK(tabnum%).Width + 120 + cmdCancel(tabnum%).Width + 480)
    cmdCancel(tabnum%).Left = cmdOK(tabnum%).Left + cmdOK(tabnum%).Width + 120
    cmdOK(tabnum%).Default = True
    fraSetup(tabnum%).Visible = True

    ' Tab index設定
    Select Case tabnum%
    Case 0
	txtMyCall.TabIndex = 1
	txtFwdCall.TabIndex = 2
	txtMyName.TabIndex = 3
	cboTimeZone.TabIndex = 4
	cmdOK(0).TabIndex = 5
	cmdCancel(0).TabIndex = 6
	'
	tkSetup.TabStop = True
	txtMyCall.TabStop = True
	txtFwdCall.TabStop = True
	txtMyName.TabStop = True
	cboTimeZone.TabStop = True
	cmdOK(0).TabStop = True
	cmdCancel(0).TabStop = True
    Case 1
	optLink(0).TabIndex = 1
	optLink(1).TabIndex = 2
	cboBaud.TabIndex = 3
	optPort(0).TabIndex = 4
	optPort(1).TabIndex = 5
	optPort(2).TabIndex = 6
	optPort(3).TabIndex = 7
	cboRcvTimeOut.TabIndex = 8
	cboSendTimeOut.TabIndex = 9
	txtSendDelay.TabIndex = 10
	cmdOK(1).TabIndex = 11
	cmdCancel(1).TabIndex = 12
	'
	optLink(0).Enabled = True
	optLink(1).Enabled = True
	optPort(0).Enabled = True
	optPort(1).Enabled = True
	optPort(2).Enabled = True
	optPort(3).Enabled = True
	'
	optLink(0).TabStop = optLink(0).Value
	optLink(1).TabStop = optLink(1).Value
	cboBaud.TabStop = True
	optPort(0).TabStop = optPort(0).Value
	optPort(1).TabStop = optPort(1).Value
	optPort(2).TabStop = optPort(2).Value
	optPort(3).TabStop = optPort(3).Value
	cboRcvTimeOut.TabStop = True
	cboSendTimeOut.TabStop = True
	cmdOK(1).TabStop = True
	cmdCancel(1).TabStop = True
    Case 2
	optMode(0).TabIndex = 1
	optMode(1).TabIndex = 2
	optMode(2).TabIndex = 3
	optMode(3).TabIndex = 4
	'
	optMode(0).TabStop = optMode(0).Value
	optMode(1).TabStop = optMode(1).Value
	optMode(2).TabStop = optMode(2).Value
	optMode(3).TabStop = optMode(3).Value
	cmdOK(2).TabStop = True
	cmdCancel(2).TabStop = True
    Case 3
	txtMsgDir.TabIndex = 1
	drvMsg.TabIndex = 2
	dirMsg.TabIndex = 3
	'
	txtMsgDir.TabStop = True
	drvMsg.TabStop = True
	dirMsg.TabStop = True
	cmdOK(3).TabStop = True
	cmdCancel(3).TabStop = True
    Case 4
	txtEditor.TabIndex = 1
	cmdEditor.TabIndex = 2
	'
	txtEditor.TabStop = True
	cmdEditor.TabStop = True
	cmdOK(4).TabStop = True
	cmdCancel(4).TabStop = True
    End Select
    Exit Sub

tkSetup_ERROR:
    If Erl = 10 Then
	Resume Next
    Else
	Error Err
    End If
End Sub

