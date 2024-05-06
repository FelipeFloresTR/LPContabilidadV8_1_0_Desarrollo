Attribute VB_Name = "PrtPreview"
Option Explicit
'Public Const SRCCOPY = &HCC0020
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, _
   ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC _
   As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Function GetPrtPage(PrtObj As Object) As Object
   
   If TypeOf PrtObj Is Printer Then
      Set GetPrtPage = Printer
   Else
      Set GetPrtPage = PrtObj.LastPage

   End If
   
End Function
Public Function NewPage(PrtObj As Object) As Object
   Dim Idx As Integer
   
   If TypeOf PrtObj Is Printer Then
      Printer.NewPage
      Set NewPage = Printer
   Else
      Set NewPage = PrtObj.NewPage
   
   End If
   
End Function
Public Sub EndDoc(PrtObj As Object)
   
   If TypeOf PrtObj Is Printer Then
      Printer.EndDoc
   End If
   
   Set PrtObj = Nothing
   
End Sub
