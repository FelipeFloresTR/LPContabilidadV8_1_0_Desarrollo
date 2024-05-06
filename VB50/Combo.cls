VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Posiciones en la lista de los datos
Private Const L_TXT = 0    ' Texto
Private Const L_DATA = 1   ' Itemdata

Private Cb As Control ' Combobox o ListBox

Private NList As Integer
Private SList() As String

Public Function AddItem(ByVal ListVal As String, Optional ByVal Data As String = "0", Optional ByVal ListVal2 As String = "", Optional ByVal ListVal3 As String = "", Optional ByVal bSelected As Boolean = 0) As Integer
   Dim n As Integer

   AddItem = -1

   If Cb Is Nothing Then
      MsgBeep vbExclamation
      Debug.Print "**** Falta asignar la ComboBox a este objeto ****"
      Exit Function
   End If

   n = Cb.ListCount
   If n + 1 > UBound(SList, 2) Then
      ReDim Preserve SList(1 To NList, n + 10)
   End If

   SList(L_DATA, n) = Data
   SList(2, n) = ListVal2
   SList(3, n) = ListVal3
   
   Cb.AddItem ListVal
   
   AddItem = Cb.NewIndex
   
   Cb.ItemData(Cb.NewIndex) = n

   If bSelected Then
      Cb.ListIndex = Cb.NewIndex
   End If

End Function
Public Sub RemoveItem(ByVal Index As Integer)

   Cb.RemoveItem Index

End Sub

Private Sub Class_Initialize()

   NList = 10
   ReDim SList(1 To NList, 10)
   
   Set Cb = Nothing

End Sub

Public Property Get list(ByVal Index As Integer) As String

   list = Cb.list(Index)
   
End Property

Public Property Let list(ByVal Index As Integer, ByVal Value As String)

   Cb.list(Index) = Value

End Property
Public Property Get ItemData(Optional ByVal Index As Integer = -1) As String

   If Index = -1 Then
      If Cb.ListIndex < 0 Then
         ItemData = -1
      Else
         
         ItemData = SList(L_DATA, Cb.ItemData(Cb.ListIndex))
      End If
   Else
      ItemData = SList(L_DATA, Cb.ItemData(Index))
   End If
End Property

Public Property Let ItemData(ByVal Index As Integer, ByVal Value As String)

   SList(L_DATA, Cb.ItemData(Index)) = Value

End Property
Public Property Get List2(Optional ByVal Index As Integer = -1) As String

   If Index = -1 Then
      If Cb.ListIndex < 0 Then
         List2 = ""
      Else
         List2 = SList(2, Cb.ItemData(Cb.ListIndex))
      End If
   Else
      List2 = SList(2, Cb.ItemData(Index))
   End If
   
End Property

Public Property Let List2(ByVal Index As Integer, ByVal Value As String)

   SList(2, Cb.ItemData(Index)) = Value

End Property
Public Property Get List3(Optional ByVal Index As Integer = -1) As String

   If Index = -1 Then
      If Cb.ListIndex < 0 Then
         List3 = ""
      Else
         List3 = SList(3, Cb.ItemData(Cb.ListIndex))
      End If
   Else
      List3 = SList(3, Cb.ItemData(Index))
   End If
   
End Property

Public Property Let List3(ByVal Index As Integer, ByVal Value As String)

   SList(3, Cb.ItemData(Index)) = Value

End Property

' ColIndex 0:  Texto
' ColIndex 1:  Itemdata
' ColIndex 2:  datos adicional - List2
' ColIndex 3:  datos adicional - List3
' ColIndex 4 a 10: datos extra

Public Property Get Matrix(ByVal ColIndex As Integer, Optional ByVal Index As Integer = -1) As String

   If Index = -1 Then
      Index = Cb.ListIndex
   End If

   If Index >= 0 Then
      If ColIndex = 0 Then ' La combo
         Matrix = Cb.list(Index)
      Else
         Matrix = SList(ColIndex, Cb.ItemData(Index))
      End If
   Else
      Matrix = ""
   End If
   
End Property

Public Property Let Matrix(ByVal ColIndex As Integer, Optional ByVal Index As Integer = -1, ByVal Value As String)

   If Index = -1 Then
      Index = Cb.ListIndex
   End If

   If ColIndex = 0 Then ' La combo
      Cb.list(Index) = Value
   Else
      SList(ColIndex, Cb.ItemData(Index)) = Value
   End If
   
End Property

Public Sub SetControl(Cntrl As Control)
   Dim Tipo As String
   
   Tipo = LCase(TypeName(Cntrl))
   
   If Tipo <> "combobox" And Tipo <> "listbox" Then
      MsgBox1 "Tipo de objeto '" & Tipo & "' no soportado.", vbExclamation
      Exit Sub
   End If
      
   Set Cb = Cntrl
   
   Call Me.Clear
   
End Sub

Public Sub Clear()

   If Not Cb Is Nothing Then
      Cb.Clear
   End If
      
End Sub

Public Property Get ListCount() As Integer
   ListCount = Cb.ListCount
End Property

Public Property Get DataChanged() As Boolean
   DataChanged = Cb.DataChanged
End Property
Public Property Get SelCount() As Long
   SelCount = Cb.SelCount
End Property

Public Property Get NewIndex() As Integer

   NewIndex = Cb.NewIndex
End Property

Public Property Get ListIndex() As Integer

   ListIndex = Cb.ListIndex
End Property

Public Property Let ListIndex(ByVal Index As Integer)

   Cb.ListIndex = Index
End Property

Public Property Get Valor() As String
Attribute Valor.VB_UserMemId = 0
Attribute Valor.VB_MemberFlags = "40"

   If Cb.ListIndex < 0 Then
      Valor = ""
   Else
      Valor = Cb.list(Cb.ListIndex)
   End If

End Property

' Se supone que sirve para Combobox y Listbox
'
' idSel = -1 ==> Se selecciona el primero
'
'  Fld(0): Item Texto
'  Fld(1): Itemdata
'  Fld(2): ...
'  ...
'  Fld(NLIST): ...
'
#If DATACON > 0 Then
#If DATACON = 1 Then
Public Function FillCombo(Db As Database, ByVal Qry As String, ByVal Sel As String, Optional ByVal bFCase As Boolean = False, Optional ByVal MaxElements As Integer = -1) As Long
#Else
Public Function FillCombo(Db As Connection, ByVal Qry As String, ByVal Sel As String, Optional ByVal bFCase As Boolean = False, Optional ByVal MaxElements As Integer = -1) As Long
#End If
   Dim Rs As Recordset
   Dim i As Integer, nf As Integer, n As Integer
   Dim Txt As String
   Dim iSel As Integer

   FillCombo = 0

   If Cb Is Nothing Then
      FillCombo = -1
      MsgBeep vbExclamation
      Debug.Print "**** Falta asignar una ComboBox a este objeto ****"
      Exit Function
   End If

   Set Rs = OpenRs(Db, Qry)
   If Rs Is Nothing Then
      FillCombo = -2
      Exit Function
   End If
   
   nf = Rs.Fields.Count ' 0 1 2 3 => 4
   If nf > NList + 1 Then
      nf = NList + 1
   End If
   
   iSel = -1
   Do Until Rs.EOF
   
      If MaxElements > 0 Then
         If Cb.ListCount > MaxElements Then
            FillCombo = -3
            Exit Do
         End If
      End If
   
      Txt = vFld(Rs(0), True)
      If bFCase Then
         Txt = FCase(Txt)
      End If
      
      n = Me.AddItem(Txt, vFld(Rs(1)))
'      n = Cb.ListCount - 1  ' En este caso, no debe ser NewIndex
      n = Cb.ItemData(n)

      For i = 2 To nf - 1
         SList(i, n) = vFld(Rs(i))
      Next i

      If Sel <> "" And Sel = vFld(Rs(1)) Then
         iSel = Cb.NewIndex
'         Cb.ListIndex = Cb.NewIndex  2 nov 2010 pam
      End If

      Rs.MoveNext
   Loop
   
   Call CloseRs(Rs)

   If Sel <> "" And iSel >= 0 Then
      Cb.ListIndex = iSel
   ElseIf Val(Sel) = -1 And Cb.ListIndex < 0 And Cb.ListCount > 0 Then
      Cb.ListIndex = 0
   End If

End Function
#End If
Public Function SelItem(ByVal Data As String) As Long
   Dim i As Long

   i = FindItem(Data)
   Cb.ListIndex = i
   SelItem = i

End Function

Public Function FindItem(ByVal Data As String, Optional ByVal ColIndex As Integer = L_DATA) As Integer
   Dim i As Integer, l As Integer
   
   Data = LCase(Trim(Data))
   
   For i = 0 To Cb.ListCount - 1
      l = Cb.ItemData(i)
      'If Val(SList(L_DATA, l)) = Data Then
      If StrComp(Trim(SList(ColIndex, l)), Data, vbTextCompare) = 0 Then
         FindItem = i
         Exit Function
      End If
   Next i

   FindItem = -1
   
End Function

Public Property Get NumList() As Integer
   NumList = NList
End Property

Public Property Let NumList(ByVal CntList As Integer)

   NList = CntList
   ReDim SList(1 To NList, 10)

End Property

Public Property Get Selected(Optional ByVal Index As Integer = -1) As Boolean

   If Index = -1 Then
      If Cb.ListIndex < 0 Then
         Selected = False
      Else
         Selected = Cb.Selected(Cb.ListIndex)
      End If
   Else
      Selected = Cb.Selected(Index)
   End If

End Property

Public Property Let Selected(ByVal Index As Integer, ByVal Value As Boolean)

   If Index = -1 Then
      If Cb.ListIndex >= 0 Then
         Cb.Selected(Cb.ListIndex) = Value
      End If
   Else
      Cb.Selected(Index) = Value
   End If

End Property
Public Sub SetSel(ByVal ItemsSelected As String)
   Dim Buf As String, i As Integer
   
   If ItemsSelected = "*" Then
      Cb.Selected(0) = True
      Cb.ListIndex = 0
      Exit Sub
      
   ElseIf ItemsSelected = "" Then
      For i = 1 To Cb.ListCount - 1
         Cb.Selected(i) = False
      Next i
      Cb.ListIndex = -1
      Exit Sub
      
   End If
      
   ItemsSelected = "," & ItemsSelected & ","
   Buf = ""
   For i = 1 To Me.ListCount - 1
      If InStr(ItemsSelected, "," & Me.ItemData(i) & ",") Then
         Me.Selected(i) = True
         Me.ListIndex = i
      End If
   Next i
   
End Sub

Public Function FindText(ByVal ItemText As String) As Long
   Dim i As Integer
   
   For i = 0 To Me.ListCount - 1
      If StrComp(Me.list(i), ItemText, vbTextCompare) = 0 Then
         FindText = i
         Exit Function
      End If
   Next i
      
   FindText = -1
      
End Function

Public Property Get Locked() As Boolean
   Locked = Cb.Locked
End Property

Public Property Let Locked(ByVal bLock As Boolean)
   Cb.Locked = bLock
End Property

Public Function IsSelected(ByVal Msg As String) As Boolean
   Dim i As Integer

   i = Cb.ListIndex
   If i >= 0 Then
      IsSelected = (Len(SList(L_DATA, Cb.ItemData(i))) > 0)
   End If

   If IsSelected = False Then
      MsgBox1 Msg, vbExclamation
   End If

End Function
