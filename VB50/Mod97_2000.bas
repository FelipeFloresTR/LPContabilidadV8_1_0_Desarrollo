Attribute VB_Name = "Mod97_2000"
Option Explicit

' Funcion para convertir Access 97 a 2000
' Se requiere Microsoft DAO 3.6

Public gDb97psw As String  ' ";psw=....

' 29 dic 2014: se crea esta funcion para poder funcionar bien en la web Sintetix, porque no soporta la version 97
Public Function Mdb97_2000(Db As Database, Fn2000 As String) As Boolean
   Dim Fn1 As String, Fn2 As String
   
   On Error Resume Next
   
   Mdb97_2000 = False
   
   If Db.Version < "4.0" Then ' "3.0" => 97, "4.0" => 2000
      Fn1 = Db.Name
      Db.Close
      Set Db = Nothing
      
      Fn2 = Left(Fn1, Len(Fn1) - 4) & "_40.mdb"
      Kill Fn2
      
      Err.Clear
      Call CompactDatabase(Fn1, Fn2, dbLangGeneral, dbVersion40, gDb97psw)

      If Err.Number = 0 Then
         Fn2000 = Fn2
         Mdb97_2000 = True
      End If

   End If

End Function

