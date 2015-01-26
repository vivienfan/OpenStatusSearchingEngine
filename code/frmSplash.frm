VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3285
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   9120
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   9120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3075
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   8865
      Begin VB.Label Label2 
         Caption         =   "Searching Engine"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   30
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3600
         TabIndex        =   4
         Top             =   1320
         Width           =   5055
      End
      Begin VB.Image imgLogo 
         Height          =   2025
         Left            =   240
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblCompany 
         Caption         =   "Designed by:  Yinghui Fan, July 2013 (2.0)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5520
         TabIndex        =   2
         Top             =   2520
         Width           =   3255
      End
      Begin VB.Label lblWarning 
         Caption         =   "This application is for Commercial Department of Siemens Building Technology (Tianjin) Ltd. use only."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1440
         TabIndex        =   1
         Top             =   2760
         Width           =   7335
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "Open Status"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   30
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   2400
         TabIndex        =   3
         Top             =   480
         Width           =   3465
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
