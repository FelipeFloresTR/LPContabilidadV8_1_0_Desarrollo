Attribute VB_Name = "ModLexis"
Option Explicit


' funciones para comunicarse con el ws de LexisNexis o LegalPublishing o ThomsonR

Public Function LPGetValorMes(ByVal Tipo As String, ByVal Ano As Integer, ByVal Mes As Integer, Optional ByVal bMsg As Boolean = 1) As Double
   Dim Req As String, Resp As String, Value As Double
   Dim i As Integer, j As Integer, bFound As Boolean
   Const Method = "TraerParamXTipoMesAnio"
   Const Url = "http://servicios.legalpublishing.cl/webservice/wsparametros/service.asmx"
   'Const Url = "http://servicios3.legalpublishing.cl/webservice/wsparametros/service.asmx"         'Hugo Lillo: 23 ago 2017
   Const Tag = "Valor>"

   LPGetValorMes = -7777
   Req = BuildReqTipoMes(Method, Tipo, Ano, Mes)
   
   If Req = "" Then
      Exit Function
   End If
   
   Resp = PostWebservice(Url, "http://tempuri.org/" & Method, Req)
   
   If Left(Resp, 5) = "Error" Then
      MsgBox1 "Falló la conexión al servidor." & vbCrLf & Resp, vbExclamation
      Exit Function
   End If

   bFound = False

   i = InStr(Resp, "<" & Tag)
   If i > 0 Then
      i = i + Len(Tag) + 1
   
      j = InStr(Resp, "</" & Tag)
   
      If j > 0 Then
         Value = Val(Mid(Resp, i, j - i))
         LPGetValorMes = Value
         bFound = True
      End If
   End If

   If bFound = False Then
      MsgBox1 "Servicio no disponible por el momento, intente más tarde.", vbExclamation
   End If

End Function

Public Function LPGetValorDia(ByVal Tipo As String, ByVal Dt As Long) As Double
   Dim Req As String, Resp As String, Value As Double
   Dim i As Integer, j As Integer, bFound As Boolean
   Const Method = "TraerParamDiarioXTipoFecha"
   Const Url = "http://servicios.legalpublishing.cl/webservice/wsparametros/service.asmx"
   'http://servicios.legalpublishing.cl/Webservice/wsparametros/service.asmx
   'Const Url = "http://servicios3.legalpublishing.cl/webservice/wsparametros/service.asmx"            'Hugo Lillo: 23 ago 2017
'   Const Url = "http://servicios3.legalpublishing.cl/webservice/wsparametros/service.asmx"            'Hugo Lillo: 23 ago 2017
   Const Tag = "Valor>"


'   If Tipo = "$US" Then   10 sep 2018
   If Tipo = "$US" Or Tipo = "US$" Then
      Tipo = "DOB"
   End If
   
   LPGetValorDia = -7777
   Req = BuildReqTipoFecha(Method, Tipo, Dt)
   
   If Req = "" Then
      Exit Function
   End If
   
   Resp = PostWebservice(Url, "http://tempuri.org/" & Method, Req)
   
   If Left(Resp, 5) = "Error" Then
      MsgBox1 "Falló la conexión al servidor." & vbCrLf & Resp, vbExclamation
      Exit Function
   End If
   
   bFound = False
   
   i = InStr(Resp, "<" & Tag)
   If i > 0 Then
      i = i + Len(Tag) + 1
   
      j = InStr(Resp, "</" & Tag)
   
      If j > 0 Then
         Value = Val(Mid(Resp, i, j - i))
         LPGetValorDia = Value
         bFound = True
      End If
   End If

   If bFound = False Then
      MsgBox1 "Servicio no disponible por el momento, intente más tarde.", vbExclamation
   End If

End Function

Private Function BuildReqTipoMes(ByVal Method As String, ByVal Tipo As String, ByVal Ano As Integer, ByVal Mes As Integer) As String
   Dim Buf As String
      
   Buf = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""no""?>"
   Buf = Buf & "<SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://www.w3.org/2003/05/soap-envelope"" xmlns:soap=""http://schemas.xmlsoap.org/wsdl/soap/"" xmlns:tm=""http://microsoft.com/wsdl/mime/textMatching/"" xmlns:soapenc=""http://schemas.xmlsoap.org/soap/encoding/"" xmlns:mime=""http://schemas.xmlsoap.org/wsdl/mime/"" xmlns:tns=""http://tempuri.org/"" xmlns:s=""http://www.w3.org/2001/XMLSchema"" xmlns:soap12=""http://schemas.xmlsoap.org/wsdl/soap12/"" xmlns:http=""http://schemas.xmlsoap.org/wsdl/http/"" xmlns:wsdl=""http://schemas.xmlsoap.org/wsdl/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" >"
   Buf = Buf & "<SOAP-ENV:Body><tns:" & Method & " xmlns:tns=""http://tempuri.org/"">"
   Buf = Buf & "<tns:tipo>" & UCase(Tipo) & "</tns:tipo>"
   Buf = Buf & "<tns:mes>" & Mes & "</tns:mes>"
   Buf = Buf & "<tns:anio>" & Ano & "</tns:anio>"
   Buf = Buf & "</tns:" & Method & ">"
   Buf = Buf & "</SOAP-ENV:Body>"
   Buf = Buf & "</SOAP-ENV:Envelope>"

   BuildReqTipoMes = Buf

End Function
Public Function BuildReqActivaLicencia(ByVal Method As String, ByVal Info As String) As String
   Dim Buf As String
      
   Buf = "<?xml version=""1.0"" encoding=""utf-8""?>"
   Buf = Buf & "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
   Buf = Buf & "<soap:Body>"
   Buf = Buf & "<ActivacionLicencia xmlns=""http://tempuri.org/"">"
   Buf = Buf & "<Info>" & Info & "</Info>"
   Buf = Buf & "</ActivacionLicencia>"
   Buf = Buf & "</soap:Body>"
   Buf = Buf & "</soap:Envelope>"

   BuildReqActivaLicencia = Buf

End Function
Private Function BuildReqTipoFecha(ByVal Method As String, ByVal Tipo As String, ByVal Dt As Long) As String
   Dim Buf As String
   
   
   Buf = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""no""?>"
   Buf = Buf & "<SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://www.w3.org/2003/05/soap-envelope"" xmlns:soap=""http://schemas.xmlsoap.org/wsdl/soap/"" xmlns:tm=""http://microsoft.com/wsdl/mime/textMatching/"" xmlns:soapenc=""http://schemas.xmlsoap.org/soap/encoding/"" xmlns:mime=""http://schemas.xmlsoap.org/wsdl/mime/"" xmlns:tns=""http://tempuri.org/"" xmlns:s=""http://www.w3.org/2001/XMLSchema"" xmlns:soap12=""http://schemas.xmlsoap.org/wsdl/soap12/"" xmlns:http=""http://schemas.xmlsoap.org/wsdl/http/"" xmlns:wsdl=""http://schemas.xmlsoap.org/wsdl/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" >"
   Buf = Buf & "<SOAP-ENV:Body><tns:" & Method & " xmlns:tns=""http://tempuri.org/"">"
'   Buf = Buf & "<tns:fecha>" & Format(Dt, "yyyy-mm-dd") & "</tns:fecha>"
   Buf = Buf & "<tns:fecha>" & Format(Dt, "mm/dd/yyyy") & "</tns:fecha>"      'Hugo Lillo: 23 ago 2017
   Buf = Buf & "<tns:tipo>" & UCase(Tipo) & "</tns:tipo>"
   Buf = Buf & "</tns:" & Method & ">"
   Buf = Buf & "</SOAP-ENV:Body>"
   Buf = Buf & "</SOAP-ENV:Envelope>"

   BuildReqTipoFecha = Buf

End Function
