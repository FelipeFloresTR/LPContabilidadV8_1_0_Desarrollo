VERSION 5.00
Begin VB.Form FrmStart 
   BorderStyle     =   0  'None
   Caption         =   "LP Contabilidad"
   ClientHeight    =   8340
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8490
   ControlBox      =   0   'False
   Icon            =   "FrmStart.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   8490
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Pc_SQLServer 
      Height          =   8235
      Left            =   0
      Picture         =   "FrmStart.frx":08CA
      ScaleHeight     =   8175
      ScaleWidth      =   8415
      TabIndex        =   6
      Top             =   0
      Width           =   8475
      Begin VB.Label La_VerSQL 
         AutoSize        =   -1  'True
         BackColor       =   &H00A67300&
         BackStyle       =   0  'Transparent
         Caption         =   "V 0.00.00"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1260
         TabIndex        =   9
         Top             =   7620
         Width           =   690
      End
      Begin VB.Label Lb_VersionSQL 
         BackStyle       =   0  'Transparent
         Caption         =   "Versión 7.0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1260
         TabIndex        =   8
         Top             =   7200
         Width           =   1515
      End
      Begin VB.Label Lb_SQL 
         BackStyle       =   0  'Transparent
         Caption         =   "Versión 7.0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2820
         TabIndex        =   7
         Top             =   7200
         Width           =   1515
      End
   End
   Begin VB.PictureBox Pc_Access 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Height          =   8325
      Left            =   -60
      Picture         =   "FrmStart.frx":34332
      ScaleHeight     =   8265
      ScaleWidth      =   8490
      TabIndex        =   0
      Top             =   0
      Width           =   8550
      Begin VB.Frame Fr_Invisible 
         Caption         =   "Invisibles"
         Height          =   1335
         Left            =   5700
         TabIndex        =   1
         Top             =   5820
         Visible         =   0   'False
         Width           =   2595
         Begin VB.DriveListBox Drive1 
            Height          =   315
            Left            =   240
            TabIndex        =   2
            Top             =   480
            Width           =   1275
         End
      End
      Begin VB.Label Lb_Access 
         BackStyle       =   0  'Transparent
         Caption         =   "Versión 7.0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A67300&
         Height          =   315
         Left            =   2760
         TabIndex        =   5
         Top             =   1800
         Width           =   1515
      End
      Begin VB.Label Lb_Version 
         BackStyle       =   0  'Transparent
         Caption         =   "Versión 7.0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A67300&
         Height          =   315
         Left            =   1200
         TabIndex        =   4
         Top             =   1800
         Width           =   1515
      End
      Begin VB.Label La_Ver 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00A67300&
         BackStyle       =   0  'Transparent
         Caption         =   "V 0.00.00"
         ForeColor       =   &H00A67300&
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   6420
         Width           =   690
      End
   End
End
Attribute VB_Name = "FrmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

   If gDbType = SQL_ACCESS Then
      Pc_SQLServer.Visible = False
      Pc_Access.Visible = True
   Else
      Pc_SQLServer.Visible = True
      Pc_Access.Visible = False
   End If

   'La_Title = gLexContab
   
   La_Ver = "V " & App.Major & "." & App.Minor & "." & App.Revision
   Lb_Version = "Versión " & App.Major & "." & App.Minor
   Lb_Access = IIf(gDbType = SQL_ACCESS, "Access", "SQL Server")
   Lb_VersionSQL = "Versión " & App.Major & "." & App.Minor
   Lb_SQL = IIf(gDbType = SQL_ACCESS, "Access", "SQL Server")
   La_VerSQL = "V " & App.Major & "." & App.Minor & "." & App.Revision
   
End Sub
  
