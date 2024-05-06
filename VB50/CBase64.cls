VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBase64"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private E(64) As Byte
Private d(256) As Byte

' Toma infomación binaria y la deja en base 64
' Si oLen = 0, no separa el buffer en partes
' cbTxt sólo incluye lo que hay que codificar, no el fin de string
Public Function Encode(inTxt() As Byte, ByVal cbTxt As Integer, Optional ByVal OLen As Byte = 76) As String
   Dim i As Long, outTxt As String, o As Integer
   Dim c As Integer, x As Integer, iSizeMod As Integer
   
   outTxt = ""
   x = 0
   o = 0
   
   iSizeMod = cbTxt Mod 3

   For i = 0 To cbTxt - iSizeMod - 1
   
      If i < cbTxt Then
         c = inTxt(i)
      ElseIf x Then ' por si quedó un resto
         c = 0
      Else
         Exit For
      End If

      Select Case (i Mod 3)
      Case 0:
         outTxt = outTxt & Chr(E(c \ 4))
         x = (c And &H3) * 16
         o = o + 1
         
         If OLen > 0 Then
            If o Mod OLen = 0 Then
               outTxt = outTxt & vbCrLf
            End If
         End If
         
      Case 1:
         outTxt = outTxt & Chr(E(x Or (c \ 16)))
         x = (c And &HF) * 4
         o = o + 1
         
         If OLen > 0 Then
            If o Mod OLen = 0 Then
               outTxt = outTxt & vbCrLf
            End If
            
         End If
         
      Case 2:
         outTxt = outTxt & Chr(E(x Or (c \ 64)))
         o = o + 1
         If OLen > 0 Then
            If o Mod OLen = 0 Then
               outTxt = outTxt & vbCrLf
            End If
         End If
         
         outTxt = outTxt & Chr(E(c And &H3F))
         o = o + 1
         
         If OLen > 0 Then
            If o Mod OLen = 0 Then
               outTxt = outTxt & vbCrLf
            End If
         End If
         
         x = 0
         
      End Select
      
   Next i

   ' El resto
   Select Case iSizeMod
      Case 1
         c = inTxt(cbTxt - 1)
         outTxt = outTxt & Chr(E(c \ 4))
         x = (c And &H3) * 16
         o = o + 1

         outTxt = outTxt & Chr(E(x))
         o = o + 1

         outTxt = outTxt & "=="
         
      Case 2
         c = inTxt(cbTxt - 2)
         x = inTxt(cbTxt - 1)
               
         '-----
         ' Output 3 bytes with 1 "="
         ' Necessary to assign "=" to round output into 1 group of 4 bytes
         '-----
         outTxt = outTxt & Chr(E((c And &HFC) \ 4))
         outTxt = outTxt & Chr(E((c And &H3) * 16 Or (x And &HF0) \ 16))
         outTxt = outTxt & Chr(E((x And &HF) * 4))
         outTxt = outTxt & "="
   End Select

   Encode = outTxt
   
End Function

Public Function EncodeTxt(ByVal inTxt As String, Optional ByVal OLen As Byte = 76) As String
   Dim l As Integer, i As Integer
   
   l = Len(inTxt)
   ReDim inBytes(l) As Byte
   
   For i = 1 To l
      inBytes(i - 1) = Asc(Mid(inTxt, i, 1))
   Next i
   inBytes(l) = 0

   EncodeTxt = Me.Encode(inBytes, l, OLen)

End Function

' Toma infomación en base 64 y la decodifica
Public Function Decode(ByVal inTxt64 As String, outTxt() As Byte) As Long
   Dim i As Long, o As Long, j As Long
   Dim c As Integer, x As Integer

   Dim cbTxt As Long
   
   cbTxt = Len(inTxt64)
   o = 0

   j = 0
   For i = 0 To cbTxt - 1

      c = d(Asc(Mid(inTxt64, i + 1, 1)))

      If c <= 63 Then

         Select Case (j Mod 4)
         Case 0:
            x = c * 4
         Case 1:
            outTxt(o) = (x Or (c \ 16))
            o = o + 1
'            outTxt = outTxt & Chr(x Or (c \ 16))
            x = (c And &HF) * 16
         Case 2:
            outTxt(o) = (x Or (c \ 4))
            o = o + 1
'            outTxt = outTxt & Chr(x Or (c \ 4))
            x = (c And &H3) * 64
         Case 3:
            outTxt(o) = (x Or c)
            o = o + 1
'            outTxt = outTxt & Chr(x Or c)
         End Select
         j = j + 1
      End If
   Next i

   outTxt(o) = 0
   o = o + 1

   Decode = o

End Function

' Toma infomación en base 64 y la decodifica
Public Function DecodeTxt(ByVal inTxt64 As String, outTxt As String) As Long
   Dim i As Long, o As Long, j As Long
   Dim c As Integer, x As Integer

   Dim cbTxt As Long
   
   cbTxt = Len(inTxt64)
   o = 0
   outTxt = Space((cbTxt * 3) / 4)

   j = 0
   For i = 0 To cbTxt - 1

'      If i = 74 And W.InDesign Then
'         Beep
'      End If

      c = d(Asc(Mid(inTxt64, i + 1, 1)))

      If c <= 63 Then ' para saltarse los cambios de linea, etc

         Select Case (j Mod 4)
         Case 0:
            x = c * 4
         Case 1:
'            outTxt(o) = (x Or (c \ 16))
            o = o + 1
            Mid(outTxt, o, 1) = Chr(x Or (c \ 16))
            x = (c And &HF) * 16
         Case 2:
            o = o + 1
            Mid(outTxt, o, 1) = Chr(x Or (c \ 4))
            x = (c And &H3) * 64
         Case 3:
            o = o + 1
            Mid(outTxt, o, 1) = Chr(x Or c)
         End Select
         
         j = j + 1
      End If
   Next i

   o = o + 1
   outTxt = Left(outTxt, o)

   DecodeTxt = o

End Function

Private Sub Class_Initialize()
   Dim i As Integer
   Dim s As String
   
   s = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"

   For i = 0 To UBound(d)
      d(i) = 64
   Next i

   For i = 0 To UBound(E) - 1
   
      E(i) = Asc(Mid(s, i + 1, 1))
   
      d(E(i)) = i
   Next i

End Sub
