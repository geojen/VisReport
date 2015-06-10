VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCompose 
   Caption         =   "Report Design and Layout"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11835
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   117
   Icon            =   "frmCompose.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8550
   ScaleWidth      =   11835
   Begin TabDlg.SSTab SSTab1 
      Height          =   8175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   14420
      _Version        =   393216
      Tabs            =   15
      TabsPerRow      =   8
      TabHeight       =   520
      ForeColor       =   16711680
      TabCaption(0)   =   "Report 1"
      TabPicture(0)   =   "frmCompose.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblTitle(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblStartYr(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblEndYr(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtTitle(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cboStartYr(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cboEndYr(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdPreview(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdPref(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame1(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Report 2"
      TabPicture(1)   =   "frmCompose.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblTitle(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblStartYr(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblEndYr(1)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cboStartYr(1)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cboEndYr(1)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdPreview(1)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmdPref(1)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "txtTitle(1)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Frame1(1)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Report 3"
      TabPicture(2)   =   "frmCompose.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblTitle(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lblStartYr(2)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "lblEndYr(2)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cboStartYr(2)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "cboEndYr(2)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "cmdPreview(2)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "cmdPref(2)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "txtTitle(2)"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Frame1(2)"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).ControlCount=   9
      TabCaption(3)   =   "Report 4"
      TabPicture(3)   =   "frmCompose.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblTitle(3)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "lblStartYr(3)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "lblEndYr(3)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "cboStartYr(3)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "cboEndYr(3)"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "cmdPreview(3)"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "cmdPref(3)"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "txtTitle(3)"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "Frame1(3)"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).ControlCount=   9
      TabCaption(4)   =   "Report 5"
      TabPicture(4)   =   "frmCompose.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "lblTitle(4)"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "lblStartYr(4)"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "lblEndYr(4)"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "cboStartYr(4)"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "cboEndYr(4)"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "cmdPreview(4)"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "cmdPref(4)"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "txtTitle(4)"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "Frame1(4)"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).ControlCount=   9
      TabCaption(5)   =   "Report 6"
      TabPicture(5)   =   "frmCompose.frx":04CE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "lblTitle(5)"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "lblStartYr(5)"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "lblEndYr(5)"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "cboStartYr(5)"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "cboEndYr(5)"
      Tab(5).Control(4).Enabled=   0   'False
      Tab(5).Control(5)=   "cmdPreview(5)"
      Tab(5).Control(5).Enabled=   0   'False
      Tab(5).Control(6)=   "cmdPref(5)"
      Tab(5).Control(6).Enabled=   0   'False
      Tab(5).Control(7)=   "txtTitle(5)"
      Tab(5).Control(7).Enabled=   0   'False
      Tab(5).Control(8)=   "Frame1(5)"
      Tab(5).Control(8).Enabled=   0   'False
      Tab(5).ControlCount=   9
      TabCaption(6)   =   "Report 7"
      TabPicture(6)   =   "frmCompose.frx":04EA
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "lblTitle(6)"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "lblStartYr(6)"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "lblEndYr(6)"
      Tab(6).Control(2).Enabled=   0   'False
      Tab(6).Control(3)=   "cboStartYr(6)"
      Tab(6).Control(3).Enabled=   0   'False
      Tab(6).Control(4)=   "cboEndYr(6)"
      Tab(6).Control(4).Enabled=   0   'False
      Tab(6).Control(5)=   "cmdPreview(6)"
      Tab(6).Control(5).Enabled=   0   'False
      Tab(6).Control(6)=   "cmdPref(6)"
      Tab(6).Control(6).Enabled=   0   'False
      Tab(6).Control(7)=   "txtTitle(6)"
      Tab(6).Control(7).Enabled=   0   'False
      Tab(6).Control(8)=   "Frame1(6)"
      Tab(6).Control(8).Enabled=   0   'False
      Tab(6).ControlCount=   9
      TabCaption(7)   =   "Report 8"
      TabPicture(7)   =   "frmCompose.frx":0506
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "lblTitle(7)"
      Tab(7).Control(0).Enabled=   0   'False
      Tab(7).Control(1)=   "lblStartYr(7)"
      Tab(7).Control(1).Enabled=   0   'False
      Tab(7).Control(2)=   "lblEndYr(7)"
      Tab(7).Control(2).Enabled=   0   'False
      Tab(7).Control(3)=   "cboStartYr(7)"
      Tab(7).Control(3).Enabled=   0   'False
      Tab(7).Control(4)=   "cboEndYr(7)"
      Tab(7).Control(4).Enabled=   0   'False
      Tab(7).Control(5)=   "cmdPreview(7)"
      Tab(7).Control(5).Enabled=   0   'False
      Tab(7).Control(6)=   "cmdPref(7)"
      Tab(7).Control(6).Enabled=   0   'False
      Tab(7).Control(7)=   "txtTitle(7)"
      Tab(7).Control(7).Enabled=   0   'False
      Tab(7).Control(8)=   "Frame1(7)"
      Tab(7).Control(8).Enabled=   0   'False
      Tab(7).ControlCount=   9
      TabCaption(8)   =   "Report 9"
      TabPicture(8)   =   "frmCompose.frx":0522
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "lblTitle(8)"
      Tab(8).Control(0).Enabled=   0   'False
      Tab(8).Control(1)=   "lblStartYr(8)"
      Tab(8).Control(1).Enabled=   0   'False
      Tab(8).Control(2)=   "lblEndYr(8)"
      Tab(8).Control(2).Enabled=   0   'False
      Tab(8).Control(3)=   "cboStartYr(8)"
      Tab(8).Control(3).Enabled=   0   'False
      Tab(8).Control(4)=   "cboEndYr(8)"
      Tab(8).Control(4).Enabled=   0   'False
      Tab(8).Control(5)=   "cmdPreview(8)"
      Tab(8).Control(5).Enabled=   0   'False
      Tab(8).Control(6)=   "cmdPref(8)"
      Tab(8).Control(6).Enabled=   0   'False
      Tab(8).Control(7)=   "txtTitle(8)"
      Tab(8).Control(7).Enabled=   0   'False
      Tab(8).Control(8)=   "Frame1(8)"
      Tab(8).Control(8).Enabled=   0   'False
      Tab(8).ControlCount=   9
      TabCaption(9)   =   "Report 10"
      TabPicture(9)   =   "frmCompose.frx":053E
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "lblTitle(9)"
      Tab(9).Control(0).Enabled=   0   'False
      Tab(9).Control(1)=   "lblStartYr(9)"
      Tab(9).Control(1).Enabled=   0   'False
      Tab(9).Control(2)=   "lblEndYr(9)"
      Tab(9).Control(2).Enabled=   0   'False
      Tab(9).Control(3)=   "cboStartYr(9)"
      Tab(9).Control(3).Enabled=   0   'False
      Tab(9).Control(4)=   "cboEndYr(9)"
      Tab(9).Control(4).Enabled=   0   'False
      Tab(9).Control(5)=   "cmdPreview(9)"
      Tab(9).Control(5).Enabled=   0   'False
      Tab(9).Control(6)=   "cmdPref(9)"
      Tab(9).Control(6).Enabled=   0   'False
      Tab(9).Control(7)=   "txtTitle(9)"
      Tab(9).Control(7).Enabled=   0   'False
      Tab(9).Control(8)=   "Frame1(9)"
      Tab(9).Control(8).Enabled=   0   'False
      Tab(9).ControlCount=   9
      TabCaption(10)  =   "Report 11"
      TabPicture(10)  =   "frmCompose.frx":055A
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "lblTitle(10)"
      Tab(10).Control(0).Enabled=   0   'False
      Tab(10).Control(1)=   "lblStartYr(10)"
      Tab(10).Control(1).Enabled=   0   'False
      Tab(10).Control(2)=   "lblEndYr(10)"
      Tab(10).Control(2).Enabled=   0   'False
      Tab(10).Control(3)=   "cboStartYr(10)"
      Tab(10).Control(3).Enabled=   0   'False
      Tab(10).Control(4)=   "cboEndYr(10)"
      Tab(10).Control(4).Enabled=   0   'False
      Tab(10).Control(5)=   "cmdPreview(10)"
      Tab(10).Control(5).Enabled=   0   'False
      Tab(10).Control(6)=   "cmdPref(10)"
      Tab(10).Control(6).Enabled=   0   'False
      Tab(10).Control(7)=   "txtTitle(10)"
      Tab(10).Control(7).Enabled=   0   'False
      Tab(10).Control(8)=   "Frame1(10)"
      Tab(10).Control(8).Enabled=   0   'False
      Tab(10).ControlCount=   9
      TabCaption(11)  =   "Report 12"
      TabPicture(11)  =   "frmCompose.frx":0576
      Tab(11).ControlEnabled=   0   'False
      Tab(11).Control(0)=   "lblTitle(11)"
      Tab(11).Control(0).Enabled=   0   'False
      Tab(11).Control(1)=   "lblStartYr(11)"
      Tab(11).Control(1).Enabled=   0   'False
      Tab(11).Control(2)=   "lblEndYr(11)"
      Tab(11).Control(2).Enabled=   0   'False
      Tab(11).Control(3)=   "cboStartYr(11)"
      Tab(11).Control(3).Enabled=   0   'False
      Tab(11).Control(4)=   "cboEndYr(11)"
      Tab(11).Control(4).Enabled=   0   'False
      Tab(11).Control(5)=   "cmdPreview(11)"
      Tab(11).Control(5).Enabled=   0   'False
      Tab(11).Control(6)=   "cmdPref(11)"
      Tab(11).Control(6).Enabled=   0   'False
      Tab(11).Control(7)=   "txtTitle(11)"
      Tab(11).Control(7).Enabled=   0   'False
      Tab(11).Control(8)=   "Frame1(11)"
      Tab(11).Control(8).Enabled=   0   'False
      Tab(11).ControlCount=   9
      TabCaption(12)  =   "Report 13"
      TabPicture(12)  =   "frmCompose.frx":0592
      Tab(12).ControlEnabled=   0   'False
      Tab(12).Control(0)=   "lblTitle(12)"
      Tab(12).Control(0).Enabled=   0   'False
      Tab(12).Control(1)=   "lblStartYr(12)"
      Tab(12).Control(1).Enabled=   0   'False
      Tab(12).Control(2)=   "lblEndYr(12)"
      Tab(12).Control(2).Enabled=   0   'False
      Tab(12).Control(3)=   "cboStartYr(12)"
      Tab(12).Control(3).Enabled=   0   'False
      Tab(12).Control(4)=   "cboEndYr(12)"
      Tab(12).Control(4).Enabled=   0   'False
      Tab(12).Control(5)=   "cmdPreview(12)"
      Tab(12).Control(5).Enabled=   0   'False
      Tab(12).Control(6)=   "cmdPref(12)"
      Tab(12).Control(6).Enabled=   0   'False
      Tab(12).Control(7)=   "txtTitle(12)"
      Tab(12).Control(7).Enabled=   0   'False
      Tab(12).Control(8)=   "Frame1(12)"
      Tab(12).Control(8).Enabled=   0   'False
      Tab(12).ControlCount=   9
      TabCaption(13)  =   "Report 14"
      TabPicture(13)  =   "frmCompose.frx":05AE
      Tab(13).ControlEnabled=   0   'False
      Tab(13).Control(0)=   "lblTitle(13)"
      Tab(13).Control(0).Enabled=   0   'False
      Tab(13).Control(1)=   "lblStartYr(13)"
      Tab(13).Control(1).Enabled=   0   'False
      Tab(13).Control(2)=   "lblEndYr(13)"
      Tab(13).Control(2).Enabled=   0   'False
      Tab(13).Control(3)=   "cboStartYr(13)"
      Tab(13).Control(3).Enabled=   0   'False
      Tab(13).Control(4)=   "cboEndYr(13)"
      Tab(13).Control(4).Enabled=   0   'False
      Tab(13).Control(5)=   "cmdPreview(13)"
      Tab(13).Control(5).Enabled=   0   'False
      Tab(13).Control(6)=   "cmdPref(13)"
      Tab(13).Control(6).Enabled=   0   'False
      Tab(13).Control(7)=   "txtTitle(13)"
      Tab(13).Control(7).Enabled=   0   'False
      Tab(13).Control(8)=   "Frame1(13)"
      Tab(13).Control(8).Enabled=   0   'False
      Tab(13).ControlCount=   9
      TabCaption(14)  =   "Report 15"
      TabPicture(14)  =   "frmCompose.frx":05CA
      Tab(14).ControlEnabled=   0   'False
      Tab(14).Control(0)=   "lblTitle(14)"
      Tab(14).Control(0).Enabled=   0   'False
      Tab(14).Control(1)=   "lblStartYr(14)"
      Tab(14).Control(1).Enabled=   0   'False
      Tab(14).Control(2)=   "lblEndYr(14)"
      Tab(14).Control(2).Enabled=   0   'False
      Tab(14).Control(3)=   "cboStartYr(14)"
      Tab(14).Control(3).Enabled=   0   'False
      Tab(14).Control(4)=   "cboEndYr(14)"
      Tab(14).Control(4).Enabled=   0   'False
      Tab(14).Control(5)=   "cmdPreview(14)"
      Tab(14).Control(5).Enabled=   0   'False
      Tab(14).Control(6)=   "cmdPref(14)"
      Tab(14).Control(6).Enabled=   0   'False
      Tab(14).Control(7)=   "txtTitle(14)"
      Tab(14).Control(7).Enabled=   0   'False
      Tab(14).Control(8)=   "Frame1(14)"
      Tab(14).Control(8).Enabled=   0   'False
      Tab(14).ControlCount=   9
      Begin VB.Frame Frame1 
         Height          =   6255
         Index           =   14
         Left            =   -74760
         TabIndex        =   205
         Top             =   1680
         Width           =   11055
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear Row"
            Height          =   375
            Index           =   14
            Left            =   120
            TabIndex        =   209
            Top             =   5760
            Width           =   1935
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "Delete Row"
            Height          =   375
            Index           =   14
            Left            =   2880
            TabIndex        =   208
            Top             =   5760
            Width           =   1935
         End
         Begin VB.CommandButton cmdBatch 
            Caption         =   "Batch Edit"
            Height          =   375
            HelpContextID   =   118
            Index           =   14
            Left            =   9000
            TabIndex        =   207
            Top             =   5760
            Width           =   1935
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "Edit Individually"
            Height          =   375
            HelpContextID   =   119
            Index           =   14
            Left            =   6240
            TabIndex        =   206
            Top             =   5760
            Width           =   1935
         End
         Begin MSFlexGridLib.MSFlexGrid grdReport 
            Height          =   5415
            Index           =   14
            Left            =   120
            TabIndex        =   210
            Top             =   240
            Width           =   10815
            _ExtentX        =   19076
            _ExtentY        =   9551
            _Version        =   393216
            ForeColor       =   16711680
            BackColorSel    =   12648447
            ForeColorSel    =   16711680
            SelectionMode   =   1
            AllowUserResizing=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame Frame1 
         Height          =   6255
         Index           =   13
         Left            =   -74760
         TabIndex        =   199
         Top             =   1680
         Width           =   11055
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear Row"
            Height          =   375
            Index           =   13
            Left            =   120
            TabIndex        =   203
            Top             =   5760
            Width           =   1935
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "Delete Row"
            Height          =   375
            Index           =   13
            Left            =   2880
            TabIndex        =   202
            Top             =   5760
            Width           =   1935
         End
         Begin VB.CommandButton cmdBatch 
            Caption         =   "Batch Edit"
            Height          =   375
            HelpContextID   =   118
            Index           =   13
            Left            =   9000
            TabIndex        =   201
            Top             =   5760
            Width           =   1935
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "Edit Individually"
            Height          =   375
            HelpContextID   =   119
            Index           =   13
            Left            =   6240
            TabIndex        =   200
            Top             =   5760
            Width           =   1935
         End
         Begin MSFlexGridLib.MSFlexGrid grdReport 
            Height          =   5415
            Index           =   13
            Left            =   120
            TabIndex        =   204
            Top             =   240
            Width           =   10815
            _ExtentX        =   19076
            _ExtentY        =   9551
            _Version        =   393216
            ForeColor       =   16711680
            BackColorSel    =   12648447
            ForeColorSel    =   16711680
            SelectionMode   =   1
            AllowUserResizing=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame Frame1 
         Height          =   6255
         Index           =   12
         Left            =   -74760
         TabIndex        =   193
         Top             =   1680
         Width           =   11055
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear Row"
            Height          =   375
            Index           =   12
            Left            =   120
            TabIndex        =   197
            Top             =   5760
            Width           =   1935
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "Delete Row"
            Height          =   375
            Index           =   12
            Left            =   2880
            TabIndex        =   196
            Top             =   5760
            Width           =   1935
         End
         Begin VB.CommandButton cmdBatch 
            Caption         =   "Batch Edit"
            Height          =   375
            HelpContextID   =   118
            Index           =   12
            Left            =   9000
            TabIndex        =   195
            Top             =   5760
            Width           =   1935
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "Edit Individually"
            Height          =   375
            HelpContextID   =   119
            Index           =   12
            Left            =   6240
            TabIndex        =   194
            Top             =   5760
            Width           =   1935
         End
         Begin MSFlexGridLib.MSFlexGrid grdReport 
            Height          =   5415
            Index           =   12
            Left            =   120
            TabIndex        =   198
            Top             =   240
            Width           =   10815
            _ExtentX        =   19076
            _ExtentY        =   9551
            _Version        =   393216
            ForeColor       =   16711680
            BackColorSel    =   12648447
            ForeColorSel    =   16711680
            SelectionMode   =   1
            AllowUserResizing=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame Frame1 
         Height          =   6255
         Index           =   11
         Left            =   -74760
         TabIndex        =   187
         Top             =   1680
         Width           =   11055
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear Row"
            Height          =   375
            Index           =   11
            Left            =   120
            TabIndex        =   191
            Top             =   5760
            Width           =   1935
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "Delete Row"
            Height          =   375
            Index           =   11
            Left            =   2880
            TabIndex        =   190
            Top             =   5760
            Width           =   1935
         End
         Begin VB.CommandButton cmdBatch 
            Caption         =   "Batch Edit"
            Height          =   375
            HelpContextID   =   118
            Index           =   11
            Left            =   9000
            TabIndex        =   189
            Top             =   5760
            Width           =   1935
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "Edit Individually"
            Height          =   375
            HelpContextID   =   119
            Index           =   11
            Left            =   6240
            TabIndex        =   188
            Top             =   5760
            Width           =   1935
         End
         Begin MSFlexGridLib.MSFlexGrid grdReport 
            Height          =   5415
            Index           =   11
            Left            =   120
            TabIndex        =   192
            Top             =   240
            Width           =   10815
            _ExtentX        =   19076
            _ExtentY        =   9551
            _Version        =   393216
            ForeColor       =   16711680
            BackColorSel    =   12648447
            ForeColorSel    =   16711680
            SelectionMode   =   1
            AllowUserResizing=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame Frame1 
         Height          =   6255
         Index           =   10
         Left            =   -74760
         TabIndex        =   181
         Top             =   1680
         Width           =   11055
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear Row"
            Height          =   375
            Index           =   10
            Left            =   120
            TabIndex        =   185
            Top             =   5760
            Width           =   1935
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "Delete Row"
            Height          =   375
            Index           =   10
            Left            =   2880
            TabIndex        =   184
            Top             =   5760
            Width           =   1935
         End
         Begin VB.CommandButton cmdBatch 
            Caption         =   "Batch Edit"
            Height          =   375
            HelpContextID   =   118
            Index           =   10
            Left            =   9000
            TabIndex        =   183
            Top             =   5760
            Width           =   1935
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "Edit Individually"
            Height          =   375
            HelpContextID   =   119
            Index           =   10
            Left            =   6240
            TabIndex        =   182
            Top             =   5760
            Width           =   1935
         End
         Begin MSFlexGridLib.MSFlexGrid grdReport 
            Height          =   5415
            Index           =   10
            Left            =   120
            TabIndex        =   186
            Top             =   240
            Width           =   10815
            _ExtentX        =   19076
            _ExtentY        =   9551
            _Version        =   393216
            ForeColor       =   16711680
            BackColorSel    =   12648447
            ForeColorSel    =   16711680
            SelectionMode   =   1
            AllowUserResizing=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame Frame1 
         Height          =   6255
         Index           =   9
         Left            =   -74760
         TabIndex        =   175
         Top             =   1680
         Width           =   11055
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear Row"
            Height          =   375
            Index           =   9
            Left            =   120
            TabIndex        =   179
            Top             =   5760
            Width           =   1935
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "Delete Row"
            Height          =   375
            Index           =   9
            Left            =   2880
            TabIndex        =   178
            Top             =   5760
            Width           =   1935
         End
         Begin VB.CommandButton cmdBatch 
            Caption         =   "Batch Edit"
            Height          =   375
            HelpContextID   =   118
            Index           =   9
            Left            =   9000
            TabIndex        =   177
            Top             =   5760
            Width           =   1935
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "Edit Individually"
            Height          =   375
            HelpContextID   =   119
            Index           =   9
            Left            =   6240
            TabIndex        =   176
            Top             =   5760
            Width           =   1935
         End
         Begin MSFlexGridLib.MSFlexGrid grdReport 
            Height          =   5415
            Index           =   9
            Left            =   120
            TabIndex        =   180
            Top             =   240
            Width           =   10815
            _ExtentX        =   19076
            _ExtentY        =   9551
            _Version        =   393216
            ForeColor       =   16711680
            BackColorSel    =   12648447
            ForeColorSel    =   16711680
            SelectionMode   =   1
            AllowUserResizing=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame Frame1 
         Height          =   6255
         Index           =   8
         Left            =   -74760
         TabIndex        =   169
         Top             =   1680
         Width           =   11055
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear Row"
            Height          =   375
            Index           =   8
            Left            =   120
            TabIndex        =   173
            Top             =   5760
            Width           =   1935
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "Delete Row"
            Height          =   375
            Index           =   8
            Left            =   2880
            TabIndex        =   172
            Top             =   5760
            Width           =   1935
         End
         Begin VB.CommandButton cmdBatch 
            Caption         =   "Batch Edit"
            Height          =   375
            HelpContextID   =   118
            Index           =   8
            Left            =   9000
            TabIndex        =   171
            Top             =   5760
            Width           =   1935
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "Edit Individually"
            Height          =   375
            HelpContextID   =   119
            Index           =   8
            Left            =   6240
            TabIndex        =   170
            Top             =   5760
            Width           =   1935
         End
         Begin MSFlexGridLib.MSFlexGrid grdReport 
            Height          =   5415
            Index           =   8
            Left            =   120
            TabIndex        =   174
            Top             =   240
            Width           =   10815
            _ExtentX        =   19076
            _ExtentY        =   9551
            _Version        =   393216
            ForeColor       =   16711680
            BackColorSel    =   12648447
            ForeColorSel    =   16711680
            SelectionMode   =   1
            AllowUserResizing=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame Frame1 
         Height          =   6255
         Index           =   7
         Left            =   -74760
         TabIndex        =   163
         Top             =   1680
         Width           =   11055
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear Row"
            Height          =   375
            Index           =   7
            Left            =   120
            TabIndex        =   167
            Top             =   5760
            Width           =   1935
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "Delete Row"
            Height          =   375
            Index           =   7
            Left            =   2880
            TabIndex        =   166
            Top             =   5760
            Width           =   1935
         End
         Begin VB.CommandButton cmdBatch 
            Caption         =   "Batch Edit"
            Height          =   375
            HelpContextID   =   118
            Index           =   7
            Left            =   9000
            TabIndex        =   165
            Top             =   5760
            Width           =   1935
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "Edit Individually"
            Height          =   375
            HelpContextID   =   119
            Index           =   7
            Left            =   6240
            TabIndex        =   164
            Top             =   5760
            Width           =   1935
         End
         Begin MSFlexGridLib.MSFlexGrid grdReport 
            Height          =   5415
            Index           =   7
            Left            =   120
            TabIndex        =   168
            Top             =   240
            Width           =   10815
            _ExtentX        =   19076
            _ExtentY        =   9551
            _Version        =   393216
            ForeColor       =   16711680
            BackColorSel    =   12648447
            ForeColorSel    =   16711680
            SelectionMode   =   1
            AllowUserResizing=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame Frame1 
         Height          =   6255
         Index           =   6
         Left            =   -74760
         TabIndex        =   157
         Top             =   1680
         Width           =   11055
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear Row"
            Height          =   375
            Index           =   6
            Left            =   120
            TabIndex        =   161
            Top             =   5760
            Width           =   1935
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "Delete Row"
            Height          =   375
            Index           =   6
            Left            =   2880
            TabIndex        =   160
            Top             =   5760
            Width           =   1935
         End
         Begin VB.CommandButton cmdBatch 
            Caption         =   "Batch Edit"
            Height          =   375
            HelpContextID   =   118
            Index           =   6
            Left            =   9000
            TabIndex        =   159
            Top             =   5760
            Width           =   1935
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "Edit Individually"
            Height          =   375
            HelpContextID   =   119
            Index           =   6
            Left            =   6240
            TabIndex        =   158
            Top             =   5760
            Width           =   1935
         End
         Begin MSFlexGridLib.MSFlexGrid grdReport 
            Height          =   5415
            Index           =   6
            Left            =   120
            TabIndex        =   162
            Top             =   240
            Width           =   10815
            _ExtentX        =   19076
            _ExtentY        =   9551
            _Version        =   393216
            ForeColor       =   16711680
            BackColorSel    =   12648447
            ForeColorSel    =   16711680
            SelectionMode   =   1
            AllowUserResizing=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame Frame1 
         Height          =   6255
         Index           =   5
         Left            =   -74760
         TabIndex        =   151
         Top             =   1680
         Width           =   11055
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear Row"
            Height          =   375
            Index           =   5
            Left            =   120
            TabIndex        =   155
            Top             =   5760
            Width           =   1935
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "Delete Row"
            Height          =   375
            Index           =   5
            Left            =   2880
            TabIndex        =   154
            Top             =   5760
            Width           =   1935
         End
         Begin VB.CommandButton cmdBatch 
            Caption         =   "Batch Edit"
            Height          =   375
            HelpContextID   =   118
            Index           =   5
            Left            =   9000
            TabIndex        =   153
            Top             =   5760
            Width           =   1935
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "Edit Individually"
            Height          =   375
            HelpContextID   =   119
            Index           =   5
            Left            =   6240
            TabIndex        =   152
            Top             =   5760
            Width           =   1935
         End
         Begin MSFlexGridLib.MSFlexGrid grdReport 
            Height          =   5415
            Index           =   5
            Left            =   120
            TabIndex        =   156
            Top             =   240
            Width           =   10815
            _ExtentX        =   19076
            _ExtentY        =   9551
            _Version        =   393216
            ForeColor       =   16711680
            BackColorSel    =   12648447
            ForeColorSel    =   16711680
            SelectionMode   =   1
            AllowUserResizing=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame Frame1 
         Height          =   6255
         Index           =   4
         Left            =   -74760
         TabIndex        =   145
         Top             =   1680
         Width           =   11055
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear Row"
            Height          =   375
            Index           =   4
            Left            =   120
            TabIndex        =   149
            Top             =   5760
            Width           =   1935
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "Delete Row"
            Height          =   375
            Index           =   4
            Left            =   2880
            TabIndex        =   148
            Top             =   5760
            Width           =   1935
         End
         Begin VB.CommandButton cmdBatch 
            Caption         =   "Batch Edit"
            Height          =   375
            HelpContextID   =   118
            Index           =   4
            Left            =   9000
            TabIndex        =   147
            Top             =   5760
            Width           =   1935
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "Edit Individually"
            Height          =   375
            HelpContextID   =   119
            Index           =   4
            Left            =   6240
            TabIndex        =   146
            Top             =   5760
            Width           =   1935
         End
         Begin MSFlexGridLib.MSFlexGrid grdReport 
            Height          =   5415
            Index           =   4
            Left            =   120
            TabIndex        =   150
            Top             =   240
            Width           =   10815
            _ExtentX        =   19076
            _ExtentY        =   9551
            _Version        =   393216
            ForeColor       =   16711680
            BackColorSel    =   12648447
            ForeColorSel    =   16711680
            SelectionMode   =   1
            AllowUserResizing=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame Frame1 
         Height          =   6255
         Index           =   3
         Left            =   -74760
         TabIndex        =   139
         Top             =   1680
         Width           =   11055
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear Row"
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   143
            Top             =   5760
            Width           =   1935
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "Delete Row"
            Height          =   375
            Index           =   3
            Left            =   2880
            TabIndex        =   142
            Top             =   5760
            Width           =   1935
         End
         Begin VB.CommandButton cmdBatch 
            Caption         =   "Batch Edit"
            Height          =   375
            HelpContextID   =   118
            Index           =   3
            Left            =   9000
            TabIndex        =   141
            Top             =   5760
            Width           =   1935
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "Edit Individually"
            Height          =   375
            HelpContextID   =   119
            Index           =   3
            Left            =   6240
            TabIndex        =   140
            Top             =   5760
            Width           =   1935
         End
         Begin MSFlexGridLib.MSFlexGrid grdReport 
            Height          =   5415
            Index           =   3
            Left            =   120
            TabIndex        =   144
            Top             =   240
            Width           =   10815
            _ExtentX        =   19076
            _ExtentY        =   9551
            _Version        =   393216
            ForeColor       =   16711680
            BackColorSel    =   12648447
            ForeColorSel    =   16711680
            SelectionMode   =   1
            AllowUserResizing=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame Frame1 
         Height          =   6255
         Index           =   2
         Left            =   -74760
         TabIndex        =   133
         Top             =   1680
         Width           =   11055
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear Row"
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   137
            Top             =   5760
            Width           =   1935
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "Delete Row"
            Height          =   375
            Index           =   2
            Left            =   2880
            TabIndex        =   136
            Top             =   5760
            Width           =   1935
         End
         Begin VB.CommandButton cmdBatch 
            Caption         =   "Batch Edit"
            Height          =   375
            HelpContextID   =   118
            Index           =   2
            Left            =   9000
            TabIndex        =   135
            Top             =   5760
            Width           =   1935
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "Edit Individually"
            Height          =   375
            HelpContextID   =   119
            Index           =   2
            Left            =   6240
            TabIndex        =   134
            Top             =   5760
            Width           =   1935
         End
         Begin MSFlexGridLib.MSFlexGrid grdReport 
            Height          =   5415
            Index           =   2
            Left            =   120
            TabIndex        =   138
            Top             =   240
            Width           =   10815
            _ExtentX        =   19076
            _ExtentY        =   9551
            _Version        =   393216
            ForeColor       =   16711680
            BackColorSel    =   12648447
            ForeColorSel    =   16711680
            SelectionMode   =   1
            AllowUserResizing=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame Frame1 
         Height          =   6255
         Index           =   1
         Left            =   -74760
         TabIndex        =   127
         Top             =   1680
         Width           =   11055
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear Row"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   131
            Top             =   5760
            Width           =   1935
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "Delete Row"
            Height          =   375
            Index           =   1
            Left            =   2880
            TabIndex        =   130
            Top             =   5760
            Width           =   1935
         End
         Begin VB.CommandButton cmdBatch 
            Caption         =   "Batch Edit"
            Height          =   375
            HelpContextID   =   118
            Index           =   1
            Left            =   9000
            TabIndex        =   129
            Top             =   5760
            Width           =   1935
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "Edit Individually"
            Height          =   375
            HelpContextID   =   119
            Index           =   1
            Left            =   6240
            TabIndex        =   128
            Top             =   5760
            Width           =   1935
         End
         Begin MSFlexGridLib.MSFlexGrid grdReport 
            Height          =   5415
            Index           =   1
            Left            =   120
            TabIndex        =   132
            Top             =   240
            Width           =   10815
            _ExtentX        =   19076
            _ExtentY        =   9551
            _Version        =   393216
            ForeColor       =   16711680
            BackColorSel    =   12648447
            ForeColorSel    =   16711680
            SelectionMode   =   1
            AllowUserResizing=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame Frame1 
         Height          =   6255
         Index           =   0
         Left            =   240
         TabIndex        =   121
         Top             =   1680
         Width           =   11055
         Begin VB.CommandButton cmdEdit 
            Caption         =   "Edit Individually"
            Height          =   375
            HelpContextID   =   119
            Index           =   0
            Left            =   6240
            TabIndex        =   126
            Top             =   5760
            Width           =   1935
         End
         Begin VB.CommandButton cmdBatch 
            Caption         =   "Batch Edit"
            Height          =   375
            HelpContextID   =   118
            Index           =   0
            Left            =   9000
            TabIndex        =   125
            Top             =   5760
            Width           =   1935
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "Delete Row"
            Height          =   375
            Index           =   0
            Left            =   2880
            TabIndex        =   123
            Top             =   5760
            Width           =   1935
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear Row"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   122
            Top             =   5760
            Width           =   1935
         End
         Begin MSFlexGridLib.MSFlexGrid grdReport 
            Height          =   5415
            Index           =   0
            Left            =   120
            TabIndex        =   124
            Top             =   240
            Width           =   10815
            _ExtentX        =   19076
            _ExtentY        =   9551
            _Version        =   393216
            ForeColor       =   16711680
            BackColorSel    =   12648447
            ForeColorSel    =   16711680
            SelectionMode   =   1
            AllowUserResizing=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.TextBox txtTitle 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   14
         Left            =   -74160
         TabIndex        =   120
         Top             =   1320
         Width           =   6855
      End
      Begin VB.TextBox txtTitle 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   13
         Left            =   -74160
         TabIndex        =   119
         Top             =   1320
         Width           =   6855
      End
      Begin VB.TextBox txtTitle 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   12
         Left            =   -74160
         TabIndex        =   118
         Top             =   1320
         Width           =   6855
      End
      Begin VB.TextBox txtTitle 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   11
         Left            =   -74160
         TabIndex        =   117
         Top             =   1320
         Width           =   6855
      End
      Begin VB.TextBox txtTitle 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   10
         Left            =   -74160
         TabIndex        =   116
         Top             =   1320
         Width           =   6855
      End
      Begin VB.TextBox txtTitle 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   9
         Left            =   -74160
         TabIndex        =   115
         Top             =   1320
         Width           =   6855
      End
      Begin VB.TextBox txtTitle 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   8
         Left            =   -74160
         TabIndex        =   114
         Top             =   1320
         Width           =   6855
      End
      Begin VB.TextBox txtTitle 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   7
         Left            =   -74160
         TabIndex        =   113
         Top             =   1320
         Width           =   6855
      End
      Begin VB.TextBox txtTitle 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   6
         Left            =   -74160
         TabIndex        =   112
         Top             =   1320
         Width           =   6855
      End
      Begin VB.TextBox txtTitle 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   5
         Left            =   -74160
         TabIndex        =   111
         Top             =   1320
         Width           =   6855
      End
      Begin VB.TextBox txtTitle 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   4
         Left            =   -74160
         TabIndex        =   110
         Top             =   1320
         Width           =   6855
      End
      Begin VB.TextBox txtTitle 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   3
         Left            =   -74160
         TabIndex        =   109
         Top             =   1320
         Width           =   6855
      End
      Begin VB.TextBox txtTitle 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   2
         Left            =   -74160
         TabIndex        =   108
         Top             =   1320
         Width           =   6855
      End
      Begin VB.TextBox txtTitle 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   1
         Left            =   -74160
         TabIndex        =   107
         Top             =   1320
         Width           =   6855
      End
      Begin VB.CommandButton cmdPref 
         Caption         =   "Page Preferences"
         Height          =   255
         HelpContextID   =   120
         Index           =   14
         Left            =   -69240
         TabIndex        =   106
         Top             =   840
         Width           =   1935
      End
      Begin VB.CommandButton cmdPreview 
         Caption         =   "View Report"
         Height          =   375
         HelpContextID   =   115
         Index           =   14
         Left            =   -64920
         TabIndex        =   105
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox cboEndYr 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   14
         Left            =   -71160
         TabIndex        =   103
         Text            =   "Combo1"
         Top             =   840
         Width           =   1335
      End
      Begin VB.ComboBox cboStartYr 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   14
         Left            =   -73680
         TabIndex        =   101
         Text            =   "Combo1"
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton cmdPref 
         Caption         =   "Page Preferences"
         Height          =   255
         HelpContextID   =   120
         Index           =   13
         Left            =   -69240
         TabIndex        =   99
         Top             =   840
         Width           =   1935
      End
      Begin VB.CommandButton cmdPreview 
         Caption         =   "View Report"
         Height          =   375
         HelpContextID   =   115
         Index           =   13
         Left            =   -64920
         TabIndex        =   98
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox cboEndYr 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   13
         Left            =   -71160
         Style           =   2  'Dropdown List
         TabIndex        =   96
         Top             =   840
         Width           =   1335
      End
      Begin VB.ComboBox cboStartYr 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   13
         Left            =   -73680
         Style           =   2  'Dropdown List
         TabIndex        =   94
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton cmdPref 
         Caption         =   "Page Preferences"
         Height          =   255
         HelpContextID   =   120
         Index           =   12
         Left            =   -69240
         TabIndex        =   92
         Top             =   840
         Width           =   1935
      End
      Begin VB.CommandButton cmdPreview 
         Caption         =   "View Report"
         Height          =   375
         HelpContextID   =   115
         Index           =   12
         Left            =   -64920
         TabIndex        =   91
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox cboEndYr 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   12
         Left            =   -71160
         Style           =   2  'Dropdown List
         TabIndex        =   89
         Top             =   840
         Width           =   1335
      End
      Begin VB.ComboBox cboStartYr 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   12
         Left            =   -73680
         Style           =   2  'Dropdown List
         TabIndex        =   87
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton cmdPref 
         Caption         =   "Page Preferences"
         Height          =   255
         HelpContextID   =   120
         Index           =   11
         Left            =   -69240
         TabIndex        =   85
         Top             =   840
         Width           =   1935
      End
      Begin VB.CommandButton cmdPreview 
         Caption         =   "View Report"
         Height          =   375
         HelpContextID   =   115
         Index           =   11
         Left            =   -64920
         TabIndex        =   84
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox cboEndYr 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   11
         Left            =   -71160
         Style           =   2  'Dropdown List
         TabIndex        =   82
         Top             =   840
         Width           =   1335
      End
      Begin VB.ComboBox cboStartYr 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   11
         Left            =   -73680
         Style           =   2  'Dropdown List
         TabIndex        =   80
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton cmdPref 
         Caption         =   "Page Preferences"
         Height          =   255
         HelpContextID   =   120
         Index           =   10
         Left            =   -69240
         TabIndex        =   78
         Top             =   840
         Width           =   1935
      End
      Begin VB.CommandButton cmdPreview 
         Caption         =   "View Report"
         Height          =   375
         HelpContextID   =   115
         Index           =   10
         Left            =   -64920
         TabIndex        =   77
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox cboEndYr 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   10
         Left            =   -71160
         Style           =   2  'Dropdown List
         TabIndex        =   75
         Top             =   840
         Width           =   1335
      End
      Begin VB.ComboBox cboStartYr 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   10
         Left            =   -73680
         Style           =   2  'Dropdown List
         TabIndex        =   73
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton cmdPref 
         Caption         =   "Page Preferences"
         Height          =   255
         HelpContextID   =   120
         Index           =   9
         Left            =   -69240
         TabIndex        =   71
         Top             =   840
         Width           =   1935
      End
      Begin VB.CommandButton cmdPreview 
         Caption         =   "View Report"
         Height          =   375
         HelpContextID   =   115
         Index           =   9
         Left            =   -64920
         TabIndex        =   70
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox cboEndYr 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   9
         Left            =   -71160
         Style           =   2  'Dropdown List
         TabIndex        =   68
         Top             =   840
         Width           =   1335
      End
      Begin VB.ComboBox cboStartYr 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   9
         Left            =   -73680
         Style           =   2  'Dropdown List
         TabIndex        =   66
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton cmdPref 
         Caption         =   "Page Preferences"
         Height          =   255
         HelpContextID   =   120
         Index           =   8
         Left            =   -69240
         TabIndex        =   64
         Top             =   840
         Width           =   1935
      End
      Begin VB.CommandButton cmdPreview 
         Caption         =   "View Report"
         Height          =   375
         HelpContextID   =   115
         Index           =   8
         Left            =   -64920
         TabIndex        =   63
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox cboEndYr 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   8
         Left            =   -71160
         Style           =   2  'Dropdown List
         TabIndex        =   61
         Top             =   840
         Width           =   1335
      End
      Begin VB.ComboBox cboStartYr 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   8
         Left            =   -73680
         Style           =   2  'Dropdown List
         TabIndex        =   59
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton cmdPref 
         Caption         =   "Page Preferences"
         Height          =   255
         HelpContextID   =   120
         Index           =   7
         Left            =   -69240
         TabIndex        =   57
         Top             =   840
         Width           =   1935
      End
      Begin VB.CommandButton cmdPreview 
         Caption         =   "View Report"
         Height          =   375
         HelpContextID   =   115
         Index           =   7
         Left            =   -64920
         TabIndex        =   56
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox cboEndYr 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   7
         Left            =   -71160
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   840
         Width           =   1335
      End
      Begin VB.ComboBox cboStartYr 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   7
         Left            =   -73680
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton cmdPref 
         Caption         =   "Page Preferences"
         Height          =   255
         HelpContextID   =   120
         Index           =   6
         Left            =   -69240
         TabIndex        =   50
         Top             =   840
         Width           =   1935
      End
      Begin VB.CommandButton cmdPreview 
         Caption         =   "View Report"
         Height          =   375
         HelpContextID   =   115
         Index           =   6
         Left            =   -64920
         TabIndex        =   49
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox cboEndYr 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   6
         Left            =   -71160
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   840
         Width           =   1335
      End
      Begin VB.ComboBox cboStartYr 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   6
         Left            =   -73680
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton cmdPref 
         Caption         =   "Page Preferences"
         Height          =   255
         HelpContextID   =   120
         Index           =   5
         Left            =   -69240
         TabIndex        =   43
         Top             =   840
         Width           =   1935
      End
      Begin VB.CommandButton cmdPreview 
         Caption         =   "View Report"
         Height          =   375
         HelpContextID   =   115
         Index           =   5
         Left            =   -64920
         TabIndex        =   42
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox cboEndYr 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   5
         Left            =   -71160
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   840
         Width           =   1335
      End
      Begin VB.ComboBox cboStartYr 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   5
         Left            =   -73680
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton cmdPref 
         Caption         =   "Page Preferences"
         Height          =   255
         HelpContextID   =   120
         Index           =   4
         Left            =   -69240
         TabIndex        =   36
         Top             =   840
         Width           =   1935
      End
      Begin VB.CommandButton cmdPreview 
         Caption         =   "View Report"
         Height          =   375
         HelpContextID   =   115
         Index           =   4
         Left            =   -64920
         TabIndex        =   35
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox cboEndYr 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   4
         Left            =   -71160
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   840
         Width           =   1335
      End
      Begin VB.ComboBox cboStartYr 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   4
         Left            =   -73680
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton cmdPref 
         Caption         =   "Page Preferences"
         Height          =   255
         HelpContextID   =   120
         Index           =   3
         Left            =   -69240
         TabIndex        =   29
         Top             =   840
         Width           =   1935
      End
      Begin VB.CommandButton cmdPreview 
         Caption         =   "View Report"
         Height          =   375
         HelpContextID   =   115
         Index           =   3
         Left            =   -64920
         TabIndex        =   28
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox cboEndYr 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   3
         Left            =   -71160
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   840
         Width           =   1335
      End
      Begin VB.ComboBox cboStartYr 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   3
         Left            =   -73680
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton cmdPref 
         Caption         =   "Page Preferences"
         Height          =   255
         HelpContextID   =   120
         Index           =   2
         Left            =   -69240
         TabIndex        =   22
         Top             =   840
         Width           =   1935
      End
      Begin VB.CommandButton cmdPreview 
         Caption         =   "View Report"
         Height          =   375
         HelpContextID   =   115
         Index           =   2
         Left            =   -64920
         TabIndex        =   21
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox cboEndYr 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   2
         Left            =   -71160
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   840
         Width           =   1335
      End
      Begin VB.ComboBox cboStartYr 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   2
         Left            =   -73680
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton cmdPref 
         Caption         =   "Page Preferences"
         Height          =   255
         HelpContextID   =   120
         Index           =   1
         Left            =   -69240
         TabIndex        =   15
         Top             =   840
         Width           =   1935
      End
      Begin VB.CommandButton cmdPreview 
         Caption         =   "View Report"
         Height          =   375
         HelpContextID   =   115
         Index           =   1
         Left            =   -64920
         TabIndex        =   14
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox cboEndYr 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   1
         Left            =   -71160
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   840
         Width           =   1335
      End
      Begin VB.ComboBox cboStartYr 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   1
         Left            =   -73680
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton cmdPref 
         Caption         =   "Page Preferences"
         Height          =   255
         HelpContextID   =   120
         Index           =   0
         Left            =   5760
         TabIndex        =   8
         Top             =   840
         Width           =   1935
      End
      Begin VB.CommandButton cmdPreview 
         Caption         =   "View Report"
         Height          =   375
         HelpContextID   =   115
         Index           =   0
         Left            =   10080
         TabIndex        =   7
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox cboEndYr 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   0
         Left            =   3840
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   840
         Width           =   1335
      End
      Begin VB.ComboBox cboStartYr 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   0
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtTitle 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   0
         Left            =   840
         TabIndex        =   2
         Top             =   1320
         Width           =   6855
      End
      Begin VB.Label lblEndYr 
         Caption         =   "End Year"
         Height          =   255
         Index           =   14
         Left            =   -72120
         TabIndex        =   104
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblStartYr 
         Caption         =   "Start Year"
         Height          =   255
         Index           =   14
         Left            =   -74760
         TabIndex        =   102
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblTitle 
         Caption         =   "Title"
         Height          =   255
         Index           =   14
         Left            =   -74760
         TabIndex        =   100
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblEndYr 
         Caption         =   "End Year"
         Height          =   255
         Index           =   13
         Left            =   -72120
         TabIndex        =   97
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblStartYr 
         Caption         =   "Start Year"
         Height          =   255
         Index           =   13
         Left            =   -74760
         TabIndex        =   95
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblTitle 
         Caption         =   "Title"
         Height          =   255
         Index           =   13
         Left            =   -74760
         TabIndex        =   93
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblEndYr 
         Caption         =   "End Year"
         Height          =   255
         Index           =   12
         Left            =   -72120
         TabIndex        =   90
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblStartYr 
         Caption         =   "Start Year"
         Height          =   255
         Index           =   12
         Left            =   -74760
         TabIndex        =   88
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblTitle 
         Caption         =   "Title"
         Height          =   255
         Index           =   12
         Left            =   -74760
         TabIndex        =   86
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblEndYr 
         Caption         =   "End Year"
         Height          =   255
         Index           =   11
         Left            =   -72120
         TabIndex        =   83
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblStartYr 
         Caption         =   "Start Year"
         Height          =   255
         Index           =   11
         Left            =   -74760
         TabIndex        =   81
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblTitle 
         Caption         =   "Title"
         Height          =   255
         Index           =   11
         Left            =   -74760
         TabIndex        =   79
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblEndYr 
         Caption         =   "End Year"
         Height          =   255
         Index           =   10
         Left            =   -72120
         TabIndex        =   76
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblStartYr 
         Caption         =   "Start Year"
         Height          =   255
         Index           =   10
         Left            =   -74760
         TabIndex        =   74
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblTitle 
         Caption         =   "Title"
         Height          =   255
         Index           =   10
         Left            =   -74760
         TabIndex        =   72
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblEndYr 
         Caption         =   "End Year"
         Height          =   255
         Index           =   9
         Left            =   -72120
         TabIndex        =   69
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblStartYr 
         Caption         =   "Start Year"
         Height          =   255
         Index           =   9
         Left            =   -74760
         TabIndex        =   67
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblTitle 
         Caption         =   "Title"
         Height          =   255
         Index           =   9
         Left            =   -74760
         TabIndex        =   65
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblEndYr 
         Caption         =   "End Year"
         Height          =   255
         Index           =   8
         Left            =   -72120
         TabIndex        =   62
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblStartYr 
         Caption         =   "Start Year"
         Height          =   255
         Index           =   8
         Left            =   -74760
         TabIndex        =   60
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblTitle 
         Caption         =   "Title"
         Height          =   255
         Index           =   8
         Left            =   -74760
         TabIndex        =   58
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblEndYr 
         Caption         =   "End Year"
         Height          =   255
         Index           =   7
         Left            =   -72120
         TabIndex        =   55
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblStartYr 
         Caption         =   "Start Year"
         Height          =   255
         Index           =   7
         Left            =   -74760
         TabIndex        =   53
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblTitle 
         Caption         =   "Title"
         Height          =   255
         Index           =   7
         Left            =   -74760
         TabIndex        =   51
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblEndYr 
         Caption         =   "End Year"
         Height          =   255
         Index           =   6
         Left            =   -72120
         TabIndex        =   48
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblStartYr 
         Caption         =   "Start Year"
         Height          =   255
         Index           =   6
         Left            =   -74760
         TabIndex        =   46
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblTitle 
         Caption         =   "Title"
         Height          =   255
         Index           =   6
         Left            =   -74760
         TabIndex        =   44
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblEndYr 
         Caption         =   "End Year"
         Height          =   255
         Index           =   5
         Left            =   -72120
         TabIndex        =   41
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblStartYr 
         Caption         =   "Start Year"
         Height          =   255
         Index           =   5
         Left            =   -74760
         TabIndex        =   39
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblTitle 
         Caption         =   "Title"
         Height          =   255
         Index           =   5
         Left            =   -74760
         TabIndex        =   37
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblEndYr 
         Caption         =   "End Year"
         Height          =   255
         Index           =   4
         Left            =   -72120
         TabIndex        =   34
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblStartYr 
         Caption         =   "Start Year"
         Height          =   255
         Index           =   4
         Left            =   -74760
         TabIndex        =   32
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblTitle 
         Caption         =   "Title"
         Height          =   255
         Index           =   4
         Left            =   -74760
         TabIndex        =   30
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblEndYr 
         Caption         =   "End Year"
         Height          =   255
         Index           =   3
         Left            =   -72120
         TabIndex        =   27
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblStartYr 
         Caption         =   "Start Year"
         Height          =   255
         Index           =   3
         Left            =   -74760
         TabIndex        =   25
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblTitle 
         Caption         =   "Title"
         Height          =   255
         Index           =   3
         Left            =   -74760
         TabIndex        =   23
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblEndYr 
         Caption         =   "End Year"
         Height          =   255
         Index           =   2
         Left            =   -72120
         TabIndex        =   20
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblStartYr 
         Caption         =   "Start Year"
         Height          =   255
         Index           =   2
         Left            =   -74760
         TabIndex        =   18
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblTitle 
         Caption         =   "Title"
         Height          =   255
         Index           =   2
         Left            =   -74760
         TabIndex        =   16
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblEndYr 
         Caption         =   "End Year"
         Height          =   255
         Index           =   1
         Left            =   -72120
         TabIndex        =   13
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblStartYr 
         Caption         =   "Start Year"
         Height          =   255
         Index           =   1
         Left            =   -74760
         TabIndex        =   11
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblTitle 
         Caption         =   "Title"
         Height          =   255
         Index           =   1
         Left            =   -74760
         TabIndex        =   9
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblEndYr 
         Caption         =   "End Year"
         Height          =   255
         Index           =   0
         Left            =   2880
         TabIndex        =   6
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblStartYr 
         Caption         =   "Start Year"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblTitle 
         Caption         =   "Title"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   1320
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmCompose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClear_Click(Index As Integer)
Dim G As MSFlexGrid
Dim N As Integer
Dim I As Integer

Set G = grdReport(Index)
If G.Row = 0 Then Exit Sub

If G.ColSel <> G.Cols - 1 Then
    MsgBox "Please Select a Row to Clear", vbInformation, "Visual Report Designer"
    Exit Sub
End If
N = G.Row

For I = G.RowSel To N Step -1
    If I <> 0 Then ClearLine Index, I
Next I

'un-highlight grid row
G.Col = 0
G.ColSel = 0

WriteLayoutFile
HTMLFile = Left(FNLOG, Len(FNLOG) - 4) & "_rpt" & CStr(Index + 1) & ".html"
If SaveLayout Then WriteHTML HTMLFile, Index
End Sub

Private Sub cmdDelete_Click(Index As Integer)
Dim G As MSFlexGrid
Dim rtnval As Integer
Dim N As Integer
Dim I As Integer

Set G = grdReport(Index)
If G.Row = 0 Then Exit Sub

If G.ColSel <> G.Cols - 1 Then
    MsgBox "Please Select an Item to Delete", vbInformation, "Visual Report Designer"
    Exit Sub
End If
N = G.Row

'If G.textmatrix(N, 1)) <> "" Then
'    rtnval = MsgBox("Delete Selected Item?", vbQuestion + vbOKCancel, "Visual Report Designer")
'    If rtnval = vbCancel Then Exit Sub
'End If

For I = G.RowSel To N Step -1
    If I <> 0 Then DeleteFromLayout Index, I
Next I

'un-highlight grid row
'G.Col = 0
'G.ColSel = 0

WriteLayoutFile
HTMLFile = Left(FNLOG, Len(FNLOG) - 4) & "_rpt" & CStr(Index + 1) & ".html"
If SaveLayout Then WriteHTML HTMLFile, Index

End Sub

Private Sub cmdEdit_Click(Index As Integer)
Dim G As MSFlexGrid
Dim I As Integer

CancelFlag = False

Set G = grdReport(Index)
If G.Row < 1 Then
    MsgBox "Please Select an Item", vbInformation, "Visual Report Designer"
    Exit Sub
End If
If G.ColSel < G.Cols - 1 Then
    MsgBox "Please Select an Item", vbInformation, "Visual Report Designer"
    Exit Sub
End If
If G.Row > G.RowSel Then
    MsgBox "Please Select Items by Highlighting them from Top to Bottom" & vbCrLf & _
            "(Instead of from Bottom to Top)", vbInformation, "Visual Report Designer"
    Exit Sub
End If

If G.Row = G.RowSel Then
    If G.TextMatrix(G.Row, 1) = "" Then Exit Sub
    MultiFlag = False
    MultiLastRow = 0
    InitSpecEdit Index, G.Row
Else
    MultiFlag = True
    MultiLastRow = G.RowSel
    InitSpecEdit Index, G.Row
End If

End Sub
Private Sub cmdBatch_Click(Index As Integer)
Dim G As MSFlexGrid
Dim I As Integer
Dim K As Integer

Set G = grdReport(Index)
If G.Row < 1 Then
    MsgBox "Please Select an Item", vbInformation, "Visual Report Designer"
    Exit Sub
End If
If G.ColSel < G.Cols - 1 Then
    MsgBox "Please Select an Item", vbInformation, "Visual Report Designer"
    Exit Sub
End If
If G.Row > G.RowSel Then
    MsgBox "Please Select Items by Highlighting them from Top to Bottom" & vbCrLf & _
            "(Instead of from Bottom to Top)", vbInformation, "Visual Report Designer"
    Exit Sub
End If

If G.Row = G.RowSel Then
    MultiFlag = False
    MultiLastRow = 0
    InitSpecEdit Index, G.Row
Else
    K = G.RowSel - G.Row + 1
    InitSpecBatch Index, G.Row, K
End If
End Sub
Private Sub GetDataEdit(Index As Integer)
Dim iRow As Integer
Dim G As MSFlexGrid
Dim I As Integer

Set G = frmCompose.grdReport(Index)
iRow = G.Row

'get the number of items to edit and their corresponding data
NDataEdit = G.RowSel - iRow + 1
If NDataEdit > 1 Then
    ReDim DataEdit(1 To NDataEdit)
    For I = 1 To NDataEdit
        DataEdit(I).Tag = BList(Index, iRow + I - 1).Tag
        DataEdit(I).Key = BList(Index, iRow + I - 1).Key
        DataEdit(I).StartYear = BList(Index, iRow + I - 1).StartYear
        DataEdit(I).EndYear = BList(Index, iRow + I - 1).EndYear
        DataEdit(I).Type = BList(Index, iRow + I - 1).Type
        DataEdit(I).LowerCut = BList(Index, iRow + I - 1).LowerCut
        DataEdit(I).UpperCut = BList(Index, iRow + I - 1).UpperCut
        DataEdit(I).Palette = BList(Index, iRow + I - 1).Palette
        DataEdit(I).ZeroFlag = BList(Index, iRow + I - 1).ZeroFlag
    Next I
End If

End Sub

Private Sub cmdPref_Click(Index As Integer)
frmLegends.lblReport.Caption = CStr(Index + 1)
InitFrmLegends
frmLegends.Show
End Sub

Private Sub Form_Load()
    Dim I As Integer
    Dim J As Integer
    Dim G As MSFlexGrid
    
    If NYears > 0 Then
    For I = 1 To 15
        For J = 1 To NYears
            cboStartYr(I - 1).AddItem CStr(StartYear + J - 1)
            cboEndYr(I - 1).AddItem CStr(StartYear + J - 1)
        Next J
        cboStartYr(I - 1).ListIndex = 0
        cboEndYr(I - 1).ListIndex = NYears - 1
    Next I
    End If
    
    For I = 1 To 15
        Set G = grdReport(I - 1)
        G.Rows = 1
        G.Cols = 10
        G.ColWidth(0) = 500
        G.TextMatrix(0, 0) = "Line"
        G.ColWidth(1) = 5000
        G.TextMatrix(0, 1) = "Description"
        G.ColWidth(2) = 800
        G.TextMatrix(0, 2) = "Start Yr"
        G.ColWidth(3) = 800
        G.TextMatrix(0, 3) = "End Yr"
        G.ColWidth(4) = 1300
        G.TextMatrix(0, 4) = "Display Type"
        G.ColWidth(5) = 2400
        G.TextMatrix(0, 5) = "Color Palette"
        G.ColWidth(6) = 1200
        G.TextMatrix(0, 6) = "High Values"
        G.ColWidth(7) = 1100
        G.TextMatrix(0, 7) = "Cut Point 1"
        G.ColWidth(8) = 1100
        G.TextMatrix(0, 8) = "Cut Point 2"
        G.ColWidth(9) = 1300
        G.TextMatrix(0, 9) = "Zero=Missing"
    Next I
    
'    For I = 2 To 15
'        SSTab1.TabEnabled(I - 1) = False
'    Next I
        
End Sub
Private Sub cboStartYr_Click(Index As Integer)
If Not InitFlag And LayoutFlag Then
    WriteLayoutFile
    HTMLFile = Left(FNLOG, Len(FNLOG) - 4) & "_rpt" & CStr(Index + 1) & ".html"
    If SaveLayout Then WriteHTML HTMLFile, Index
End If
End Sub
Private Sub cboEndYr_Click(Index As Integer)
If Not InitFlag And LayoutFlag Then
    WriteLayoutFile
    HTMLFile = Left(FNLOG, Len(FNLOG) - 4) & "_rpt" & CStr(Index + 1) & ".html"
    If SaveLayout Then WriteHTML HTMLFile, Index
End If
End Sub
Private Sub txtTitle_Change(Index As Integer)
If Not InitFlag And LayoutFlag Then
    WriteLayoutFile
    HTMLFile = Left(FNLOG, Len(FNLOG) - 4) & "_rpt" & CStr(Index + 1) & ".html"
    If SaveLayout Then WriteHTML HTMLFile, Index
End If
End Sub
Private Sub cmdPreview_Click(Index As Integer)
Dim G As MSFlexGrid
Dim CMD As String
Dim KL As Long
    
Set G = frmCompose.grdReport(Index)
If G.Rows < 2 Then
    MsgBox "Please add data to Report " & CStr(Index + 1), vbInformation, "Visual Report Designer"
    Exit Sub
End If
If Not InitFlag And LayoutFlag Then
    WriteLayoutFile
    HTMLFile = Left(FNLOG, Len(FNLOG) - 4) & "_rpt" & CStr(Index + 1) & ".html"
    If SaveLayout Then WriteHTML HTMLFile, Index
End If

If RptViewer = "GUI" Then
    frmPreview.Show
ElseIf RptViewer = "Browser" Then
    KL = ShellExecute(frmMain.hWnd, "open", HTMLFile, 0, 0, 1)
    If KL < 32 Then
        MsgBox "Unable to Open " & HTMLFile, vbExclamation, "Visual Report Designer"
    End If
Else
    CMD = RptViewer & " " & HTMLFile
    Shell CMD, vbNormalFocus
End If

End Sub
