VERSION 5.00
Begin VB.Form FrmPrintPreview 
   Caption         =   "Form1"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10725
   Icon            =   "FrmPrintPreview.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   10725
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bt_Cerrar 
      Caption         =   "Cerrar"
      Height          =   315
      Left            =   9420
      TabIndex        =   6
      Top             =   60
      Width           =   1215
   End
   Begin VB.Frame Fr_Invisible 
      Caption         =   "No se Ve"
      Height          =   555
      Left            =   60
      TabIndex        =   13
      Top             =   6300
      Visible         =   0   'False
      Width           =   1035
      Begin VB.Image Im_Mouse 
         Height          =   315
         Left            =   60
         MouseIcon       =   "FrmPrintPreview.frx":000C
         Top             =   180
         Width           =   375
      End
      Begin VB.Image Im_Menos 
         Height          =   315
         Left            =   600
         MouseIcon       =   "FrmPrintPreview.frx":015E
         Top             =   180
         Width           =   375
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   6735
      Left            =   10440
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   780
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   60
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   7620
      Width           =   10275
   End
   Begin VB.CommandButton bt_Ini 
      Caption         =   "&Inicio"
      Height          =   615
      Left            =   60
      Picture         =   "FrmPrintPreview.frx":02B0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton bt_Sgte 
      Caption         =   "&Siguiente"
      Height          =   615
      Left            =   1980
      Picture         =   "FrmPrintPreview.frx":05BA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton bt_Atras 
      Caption         =   "&Atrás"
      Height          =   615
      Left            =   1020
      Picture         =   "FrmPrintPreview.frx":08C4
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton bt_Final 
      Caption         =   "&Final"
      Height          =   615
      Left            =   2940
      Picture         =   "FrmPrintPreview.frx":0BCE
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton bt_Prt 
      Caption         =   "&Imprimir"
      Height          =   615
      Left            =   4860
      Picture         =   "FrmPrintPreview.frx":0ED8
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   60
      Width           =   915
   End
   Begin VB.PictureBox Pc_Fondo 
      BackColor       =   &H8000000C&
      Height          =   6735
      Left            =   60
      ScaleHeight     =   6675
      ScaleWidth      =   10215
      TabIndex        =   7
      Top             =   780
      Width           =   10275
      Begin VB.PictureBox Pc_PicView 
         BackColor       =   &H80000005&
         Height          =   6555
         Left            =   60
         MousePointer    =   99  'Custom
         ScaleHeight     =   6495
         ScaleWidth      =   10035
         TabIndex        =   12
         Top             =   60
         Width           =   10095
      End
      Begin VB.PictureBox Pc_Preview 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H00000000&
         Height          =   6555
         Index           =   0
         Left            =   60
         MousePointer    =   99  'Custom
         ScaleHeight     =   6525
         ScaleWidth      =   10065
         TabIndex        =   8
         Tag             =   "False"
         Top             =   60
         Visible         =   0   'False
         Width           =   10095
      End
   End
   Begin VB.CommandButton bt_Zoom 
      Caption         =   "&Zoom"
      Height          =   615
      Left            =   3900
      Picture         =   "FrmPrintPreview.frx":1392
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   60
      Width           =   915
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5940
      TabIndex        =   9
      Top             =   60
      Width           =   1515
   End
End
Attribute VB_Name = "FrmPrintPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public PrtControl As CommandButton

Dim lPagVisible As Integer
Dim lZoom As Boolean
Dim isRectExist As Boolean
Dim isBoxExist As Boolean
Dim sWidth As Single
Dim sHeight As Single
Dim originalHeight As Single
Dim originalWidth As Single
Dim minX As Single
Dim maxX As Single
Dim minY As Single
Dim maxY As Single
Dim lCap As String
Dim PicActual As Picture

Const vScrollMax = 150
Const hScrollMax = 35
Const mZoomFactor = 1

Private Sub bt_Atras_Click()
   
   If lPagVisible > 0 Then
      lPagVisible = lPagVisible - 1
   
   End If
   Call MovPag(lPagVisible > 0, True)
   
End Sub

Private Sub bt_Final_Click()
   
   lPagVisible = Pc_Preview.Count - 1
   Call MovPag(True And Pc_Preview.Count > 1, False)
   
End Sub

Private Sub bt_Ini_Click()

   lPagVisible = 0
   Call MovPag(False, True And Pc_Preview.Count > 1)
   
End Sub

Private Sub bt_Prt_Click()

   If Not PrtControl Is Nothing Then
      Call PostClick(PrtControl, False)
   End If
   
End Sub

Private Sub bt_Sgte_Click()

   If lPagVisible < Pc_Preview.Count - 1 Then
      lPagVisible = lPagVisible + 1
      
   End If
   Call MovPag(True, lPagVisible < Pc_Preview.Count - 1)
   
End Sub

Private Sub bt_Zoom_Click()

   On Error Resume Next
   
   Set Pc_PicView = Nothing
   Pc_PicView.Picture = Pc_Preview(lPagVisible).Image
   
   Call ZoomImgInOut(Not lZoom)
   
End Sub

Private Sub Form_Activate()
  
   'Luego Seteo Pic que ocupo cuando marco ya sea en zoom in o zoom out
   Set PicActual = Pc_Preview(lPagVisible).Image
   
   'Seteo Pic View
   Set Pc_PicView.Picture = Pc_Preview(lPagVisible).Image
   
   'Inicializo variable que voy a necesitar más adelante
   label1 = "Página 1 de " & Pc_Preview.Count
   Caption = lCap
         
   lPagVisible = 0
   bt_Sgte.Enabled = Pc_Preview.Count > 1
   bt_Atras.Enabled = False
   
   VScroll1.Max = vScrollMax
   VScroll1.LargeChange = 5
   VScroll1.SmallChange = 20
   
   HScroll1.Max = hScrollMax
   HScroll1.LargeChange = 5
   HScroll1.SmallChange = 20
   
   originalHeight = Pc_Preview(lPagVisible).Height
   originalWidth = Pc_Preview(lPagVisible).Width
   
   Pc_PicView.Height = originalHeight
   Pc_PicView.Width = originalWidth
   
End Sub

Private Sub bt_Cerrar_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   label1 = "Página 1 de " & Pc_Preview.Count
   
   'seteo con el tamaño de la impresora
   Pc_Preview(0).Height = Printer.Height '* (Printer.Height / Printer.ScaleHeight)
   Pc_Preview(0).Width = Printer.Width '* (Printer.Width / Printer.ScaleWidth)
   
   Pc_Preview(0).ScaleMode = Printer.ScaleMode
   Pc_Preview(0).ScaleWidth = Printer.ScaleWidth
   Pc_Preview(0).ScaleHeight = Printer.ScaleHeight
   Pc_Preview(0).ScaleTop = Printer.ScaleTop
   Pc_Preview(0).ScaleLeft = Printer.ScaleLeft
      
   Set PrtControl = Nothing
   
End Sub
Public Function NewPage() As PictureBox
   Dim Idx As Integer
   
   Idx = Pc_Preview.Count
   Load Pc_Preview(Idx)
   Pc_Preview(Idx).Visible = False
   Set NewPage = Pc_Preview(Idx)

End Function

Public Function LastPage() As PictureBox
   Dim Idx As Integer
   
   Idx = Pc_Preview.Count - 1
   Set LastPage = Pc_Preview(Idx)

End Function

Private Sub Form_Resize()
    Dim d As Integer

   If Me.WindowState = vbMinimized Then
      Exit Sub
   End If

   d = Me.Width - 2 * (Pc_Fondo.Left + W.xFrame) - (VScroll1.Width + 50)
   If d > 1000 Then
      Pc_Fondo.Width = d
   End If
   
   d = Me.Height - Pc_Fondo.Top - 200 - W.YCaption * 2
   If d > 1000 Then
      Pc_Fondo.Height = d
   Else
      Me.Height = Pc_Fondo.Top + 1000 + W.YCaption * 2
   End If
   
   VScroll1.Left = Pc_Fondo.Width + 100
   VScroll1.Height = Pc_Fondo.Height
   
   HScroll1.Visible = Pc_Preview(lPagVisible).Image.Width > Pc_Fondo.Width
   HScroll1.Top = Pc_Fondo.Top + Pc_Fondo.Height
   HScroll1.Width = Pc_Fondo.Width
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set Pc_PicView = Nothing
   Set PicActual = Nothing
      
End Sub

Private Sub Pc_PicView_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   If Button = 1 Then
      If isRectExist Then
         Pc_PicView.Cls
         isBoxExist = False
      End If
      
      Set Pc_PicView.Picture = PicActual
      
      minX = x
      maxY = Y
      maxX = x
      minY = Y
   End If
   
End Sub

Private Sub Pc_PicView_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
   
    If Button = 1 Then
      Pc_PicView.MouseIcon = Me.Im_Mouse.MouseIcon
      
      Pc_PicView.DrawMode = 10
      Pc_PicView.Line (minX, maxY)-(maxX, minY), vbBlue, BF
      maxX = x
      minY = Y
      Pc_PicView.Line (minX, maxY)-(maxX, minY), vbBlue, BF
      Pc_PicView.DrawMode = 13
          
      x = x + Pc_PicView.Left - Pc_Fondo.Left
      Y = Y + Pc_PicView.Top + Pc_Fondo.Top
      
      If x > Pc_Fondo.Left + Pc_Fondo.Width And HScroll1.Value < HScroll1.Max Then
         HScroll1.Value = HScroll1.Value + 1
         Pc_PicView.Left = ((HScroll1.Value / 100) * ScaleWidth) * -1
      ElseIf x < Pc_Fondo.Left And HScroll1.Value > 0 Then
         HScroll1.Value = HScroll1.Value - 1
         Pc_PicView.Left = ((HScroll1.Value / 100) * ScaleWidth) * -1
      End If
       
      If Y > Pc_Fondo.Top + Pc_Fondo.Height And VScroll1.Value < VScroll1.Max Then
         VScroll1.Value = VScroll1.Value + 1
         Pc_PicView.Top = ((VScroll1.Value / 100) * ScaleHeight) * -1
      ElseIf Y < Pc_Fondo.Top And VScroll1.Value > 0 Then
         VScroll1.Value = VScroll1.Value - 1
         Pc_PicView.Top = ((VScroll1.Value / 100) * ScaleHeight) * -1
      End If
      
    End If
    
End Sub

Private Sub Pc_PicView_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    isRectExist = True
    Pc_PicView.MouseIcon = Nothing
    
End Sub

Private Sub VScroll1_Change()
   Pc_PicView.Top = ((VScroll1.Value / 100) * ScaleHeight) * -1
End Sub
Private Sub HScroll1_Change()
   Pc_PicView.Left = ((HScroll1.Value / 100) * ScaleWidth) * -1
End Sub
Private Sub ZoomImgInOut(bool As Boolean)
   
   On Error Resume Next
   
   Set PicActual = Nothing
   
   If bool Then
      'Zoom In
      sHeight = originalHeight + (originalHeight * mZoomFactor)
      sWidth = originalWidth + (originalWidth * mZoomFactor)
      
      VScroll1.Max = vScrollMax + 200
      HScroll1.Max = hScrollMax + 150
      
      Call FitPicture(Pc_PicView, "", sWidth, sHeight)
      
   Else
      'Zoom Out
      sHeight = originalHeight
      sWidth = originalWidth
      
      Pc_PicView.Width = sWidth
      Pc_PicView.Height = sHeight
      
      VScroll1.Max = vScrollMax
      HScroll1.Max = hScrollMax
   End If
   
   lZoom = bool
   
   Set PicActual = Pc_PicView.Image
   
   HScroll1.Visible = (sWidth > Pc_Fondo.Width)
   VScroll1.Visible = (sHeight > Pc_Fondo.Height)
   
End Sub
Private Sub MovPag(BoolAtras As Boolean, BoolSgte As Boolean)

   bt_Atras.Enabled = BoolAtras
   bt_Sgte.Enabled = BoolSgte
   
   Set Pc_PicView = Nothing
   Pc_PicView.Picture = Pc_Preview(lPagVisible).Image
   
   label1 = "Página " & lPagVisible + 1 & " de " & Pc_Preview.Count
   
   If lZoom Then
      Call ZoomImgInOut(lZoom)
   End If
   
End Sub
Public Function FView(Cap As String)
   lCap = Cap
   Me.Show vbModal
   
End Function
