Attribute VB_Name = "VBZipBas"

Option Explicit

'---------------------------------------------------------------
'-- Please Do Not Remove These Comments!!!
'---------------------------------------------------------------
'-- Sample VB 5 code to drive zip32.dll
'-- Contributed to the Info-ZIP project by Mike Le Voi
'--
'-- Contact me at: mlevoi@modemss.brisnet.org.au
'--
'-- Visit my home page at: http://modemss.brisnet.org.au/~mlevoi
'--
'-- Use this code at your own risk. Nothing implied or warranted
'-- to work on your machine :-)
'---------------------------------------------------------------
'--
'-- The Source Code Is Freely Available From Info-ZIP At:
'-- http://www.cdrom.com/pub/infozip/infozip.html
'--
'-- A Very Special Thanks To Mr. Mike Le Voi
'-- And Mr. Mike White Of The Info-ZIP
'-- For Letting Me Use And Modify His Orginal
'-- Visual Basic 5.0 Code! Thank You Mike Le Voi.
'---------------------------------------------------------------
'--
'-- Contributed To The Info-ZIP Project By Raymond L. King
'-- Modified June 21, 1998
'-- By Raymond L. King
'-- Custom Software Designers
'--
'-- Contact Me At: king@ntplx.net
'-- ICQ 434355
'-- Or Visit Our Home Page At: http://www.ntplx.net/~king
'--
'---------------------------------------------------------------
' Common Options
'   Opt.Date = vbNullString
'   Opt.fJunkDir = 0     ' 1 = Throw Away Path Names
'   Opt.fRecurse = 0     ' 1 = Recurse -R 2 = Recurse -r 2 = Most Useful :)
'   Opt.fUpdate = 0      ' 1 = Update Only If Newer
'   Opt.fFreshen = 0     ' 1 = Freshen - Overwrite Only
'   Opt.flevel = Asc(9)  ' Compression Level (0 - 9)
'   Opt.fEncrypt = 0 '1     ' Encryption = 1 For Password Else 0
'   Opt.fComment = 0 '1     ' Comment = 1 if required
'   Opt.szRootDir = "c:\"

'-- C Style argv
'-- Holds The Zip Archive Filenames
Public Type ZIPnames_t
  zFiles(0 To 99) As String
End Type

'-- Call Back "String"
Public Type ZipCBChar_t
  ch(4096) As Byte
End Type

'-- ZPOPT Is Used To Set The Options In The ZIP32.DLL
Public Type ZipOPT_t
  Date           As String ' US Date (8 Bytes Long) "12/31/98"?
  szRootDir      As String ' Root Directory Pathname (Up To 256 Bytes Long)
  szTempDir      As String ' Temp Directory Pathname (Up To 256 Bytes Long)
  fTemp          As Long   ' 1 If Temp dir Wanted, Else 0
  fSuffix        As Long   ' Include Suffixes (Not Yet Implemented!)
  fEncrypt       As Long   ' 1 If Encryption Wanted, Else 0
  fSystem        As Long   ' 1 To Include System/Hidden Files, Else 0
  fVolume        As Long   ' 1 If Storing Volume Label, Else 0
  fExtra         As Long   ' 1 If Excluding Extra Attributes, Else 0
  fNoDirEntries  As Long   ' 1 If Ignoring Directory Entries, Else 0
  fExcludeDate   As Long   ' 1 If Excluding Files Earlier Than Specified Date, Else 0
  fIncludeDate   As Long   ' 1 If Including Files Earlier Than Specified Date, Else 0
  fVerbose       As Long   ' 1 If Full Messages Wanted, Else 0
  fQuiet         As Long   ' 1 If Minimum Messages Wanted, Else 0
  fCRLF_LF       As Long   ' 1 If Translate CR/LF To LF, Else 0
  fLF_CRLF       As Long   ' 1 If Translate LF To CR/LF, Else 0
  fJunkDir       As Long   ' 1 If Junking Directory Names, Else 0
  fGrow          As Long   ' 1 If Allow Appending To Zip File, Else 0
  fForce         As Long   ' 1 If Making Entries Using DOS File Names, Else 0
  fMove          As Long   ' 1 If Deleting Files Added Or Updated, Else 0
  fDeleteEntries As Long   ' 1 If Files Passed Have To Be Deleted, Else 0
  fUpdate        As Long   ' 1 If Updating Zip File-Overwrite Only If Newer, Else 0
  fFreshen       As Long   ' 1 If Freshing Zip File-Overwrite Only, Else 0
  fJunkSFX       As Long   ' 1 If Junking SFX Prefix, Else 0
  fLatestTime    As Long   ' 1 If Setting Zip File Time To Time Of Latest File In Archive, Else 0
  fComment       As Long   ' 1 If Putting Comment In Zip File, Else 0
  fOffsets       As Long   ' 1 If Updating Archive Offsets For SFX Files, Else 0
  fPrivilege     As Long   ' 1 If Not Saving Privileges, Else 0
  fEncryption    As Long   ' Read Only Property!!!
  fRecurse       As Long   ' 1 (-r), 2 (-R) If Recursing Into Sub-Directories, Else 0
  fRepair        As Long   ' 1 = Fix Archive, 2 = Try Harder To Fix, Else 0
  flevel         As Byte   ' Compression Level - 0 = Stored 6 = Default 9 = Max
End Type

'-- This Structure Is Used For The ZIP32.DLL Function Callbacks
Public Type ZIPUSERFUNCTIONS_t
  ZDLLPrnt     As Long        ' Callback ZIP32.DLL Print Function
  ZDLLCOMMENT  As Long        ' Callback ZIP32.DLL Comment Function
  ZDLLPASSWORD As Long        ' Callback ZIP32.DLL Password Function
  ZDLLSERVICE  As Long        ' Callback ZIP32.DLL Service Function
End Type

'-- Local Declarations
'Public gZipOpt  As ZipOPT_t
'Public gZipUserFnc As ZIPUSERFUNCTIONS_t
Public gnZippedFiles As Long

'-- This Assumes ZIP32.DLL Is In Your \Windows\System Directory!
'-- (alternatively, a copy of ZIP32.DLL needs to be located in the program
'-- directory or in some other directory listed in the PATH.)
Private Declare Function ZpInit Lib "FwZip32.dll" _
  (ByRef Zipfun As ZIPUSERFUNCTIONS_t) As Long '-- Set Zip Callbacks

Private Declare Function ZpSetOptions Lib "FwZip32.dll" _
  (ByRef Opts As ZipOPT_t) As Long '-- Set Zip Options

Private Declare Function ZpGetOptions Lib "FwZip32.dll" _
  () As ZipOPT_t '-- Used To Check Encryption Flag Only

Private Declare Function ZpArchive Lib "FwZip32.dll" _
  (ByVal argc As Long, ByVal funame As String, _
   ByRef argv As ZIPnames_t) As Long '-- Real Zipping Action

'-------------------------------------------------------
'-- Public Variables For Setting The ZPOPT Structure...
'-- (WARNING!!!) You Must Set The Options That You
'-- Want The ZIP32.DLL To Do!
'-- Before Calling VBZip32!
'--
'-- NOTE: See The Above ZPOPT Structure Or The VBZip32
'--       Function, For The Meaning Of These Variables
'--       And How To Use And Set Them!!!
'-- These Parameters Must Be Set Before The Actual Call
'-- To The VBZip32 Function!
'-------------------------------------------------------
Public zDate         As String
Public zRootDir      As String
Public zTempDir      As String
Public zSuffix       As Integer
Public zEncrypt      As Integer
Public zSystem       As Integer
Public zVolume       As Integer
Public zExtra        As Integer
Public zNoDirEntries As Integer
Public zExcludeDate  As Integer
Public zIncludeDate  As Integer
Public zVerbose      As Integer
Public zQuiet        As Integer
Public zCRLF_LF      As Integer
Public zLF_CRLF      As Integer
Public zJunkDir      As Integer
Public zRecurse      As Integer
Public zGrow         As Integer
Public zForce        As Integer
Public zMove         As Integer
Public zDelEntries   As Integer
Public zUpdate       As Integer
Public zFreshen      As Integer
Public zJunkSFX      As Integer
Public zLatestTime   As Integer
Public zComment      As Integer
Public zOffsets      As Integer
Public zPrivilege    As Integer
Public zEncryption   As Integer
Public zRepair       As Integer
Public zLevel        As Integer

'-- Public Program Variables
Public zArgc         As Integer     ' Number Of Files To Zip Up
Public zZipFileName  As String      ' The Zip File Name ie: Myzip.zip
Public zZipFileNames As ZIPnames_t    ' File Names To Zip Up
Public zZipInfo      As String      ' Holds The Zip File Information

'-- Public Constants
'-- For Zip & UnZip Error Codes!
Public Const ZE_OK = 0              ' Success (No Error)
Public Const ZE_EOF = 2             ' Unexpected End Of Zip File Error
Public Const ZE_FORM = 3            ' Zip File Structure Error
Public Const ZE_MEM = 4             ' Out Of Memory Error
Public Const ZE_LOGIC = 5           ' Internal Logic Error
Public Const ZE_BIG = 6             ' Entry Too Large To Split Error
Public Const ZE_NOTE = 7            ' Invalid Comment Format Error
Public Const ZE_TEST = 8            ' Zip Test (-T) Failed Or Out Of Memory Error
Public Const ZE_ABORT = 9           ' User Interrupted Or Termination Error
Public Const ZE_TEMP = 10           ' Error Using A Temp File
Public Const ZE_READ = 11           ' Read Or Seek Error
Public Const ZE_NONE = 12           ' Nothing To Do Error
Public Const ZE_NAME = 13           ' Missing Or Empty Zip File Error
Public Const ZE_WRITE = 14          ' Error Writing To A File
Public Const ZE_CREAT = 15          ' Could't Open To Write Error
Public Const ZE_PARMS = 16          ' Bad Command Line Argument Error
Public Const ZE_OPEN = 18           ' Could Not Open A Specified File To Read Error

'-- These Functions Are For The ZIP32.DLL
'--
'-- Puts A Function Pointer In A Structure
'-- For Use With Callbacks...
Public Function FnPtr(ByVal lp As Long) As Long
    
  FnPtr = lp

End Function

'-- Callback For ZIP32.DLL - DLL Print Function
Private Function ZDLLPrnt(ByRef fname As ZipCBChar_t, ByVal x As Long) As Long
  Dim s0 As String
  Dim xx As Long
    
  ZDLLPrnt = 0
  
  Exit Function
    
  '-- Always Put This In Callback Routines!
  On Error Resume Next
    
  s0 = ""
    
  '-- Get Zip32.DLL Message For processing
  For xx = 0 To x
    If fname.ch(xx) = 0 Then
      Exit For
    Else
      s0 = s0 + Chr(fname.ch(xx))
    End If
  Next
    
  '----------------------------------------------
  '-- This Is Where The DLL Passes Back Messages
  '-- To You! You Can Change The Message Printing
  '-- Below Here!
  '----------------------------------------------
  
  '-- Display Zip File Information
  '-- zZipInfo = zZipInfo & s0
  'Form1.Print s0;
    
  'DoEvents
    
  ZDLLPrnt = 0

End Function

'-- Callback For ZIP32.DLL - DLL Service Function
Private Function ZDLLServ(ByRef mname As ZipCBChar_t, ByVal x As Long) As Long

    Dim s0 As String
    Dim xx As Long
    
    '-- Always Put This In Callback Routines!
    On Error Resume Next
    
    gnZippedFiles = gnZippedFiles + 1
    
    s0 = ""
    '-- Get Zip32.DLL Message For processing
    For xx = 0 To x
      If mname.ch(xx) = 0 Then
          Exit For
      Else
          s0 = s0 + Chr(mname.ch(xx))
      End If
    Next
    ' At this point, s0 contains the message passed from the DLL
    ' It is up to the developer to code something useful here :)
    ZDLLServ = 0 ' Setting this to 1 will abort the zip!
    
End Function

'-- Callback For ZIP32.DLL - DLL Password Function
Private Function ZDLLPass(ByRef p As ZipCBChar_t, _
  ByVal n As Long, ByRef m As ZipCBChar_t, _
  ByRef Name As ZipCBChar_t) As Integer
  
  Dim prompt     As String
  Dim xx         As Integer
  Dim szpassword As String
  
  '-- Always Put This In Callback Routines!
  On Error Resume Next
    
  ZDLLPass = 1
  
  '-- If There Is A Password Have The User Enter It!
  '-- This Can Be Changed
  szpassword = InputBox("Please Enter The Password!")
  
  '-- The User Did Not Enter A Password So Exit The Function
  If szpassword = "" Then Exit Function
  
  '-- User Entered A Password So Proccess It
  For xx = 0 To 255
    If m.ch(xx) = 0 Then
      Exit For
    Else
      prompt = prompt & Chr(m.ch(xx))
    End If
  Next
  
  For xx = 0 To n - 1
    p.ch(xx) = 0
  Next
  
  For xx = 0 To Len(szpassword) - 1
    p.ch(xx) = Asc(Mid(szpassword, xx + 1, 1))
  Next
  
  p.ch(xx) = Chr(0) ' Put Null Terminator For C
  
  ZDLLPass = 0
    
End Function

'-- Callback For ZIP32.DLL - DLL Comment Function
Private Function ZDLLComm(ByRef s1 As ZipCBChar_t) As Integer
    
    Dim xx%, szcomment$
    
    '-- Always Put This In Callback Routines!
    On Error Resume Next
    
    ZDLLComm = 1
    szcomment = InputBox("Enter the comment")
    If szcomment = "" Then Exit Function
    For xx = 0 To Len(szcomment) - 1
        s1.ch(xx) = Asc(Mid$(szcomment, xx + 1, 1))
    Next xx
    s1.ch(xx) = Chr(0) ' Put null terminator for C

End Function

'-- Main ZIP32.DLL Subroutine.
'-- This Is Where It All Happens!!!
'--
'-- (WARNING!) Do Not Change This Function!!!
'--
'-- Si las funciones en UsrFnc no se indican, se asignan unas por defecto
Public Function VBZip32(ByVal ZipFName As String, ByVal nFileNames As Integer, FileNames As ZIPnames_t, Opt As ZipOPT_t, UserFnc As ZIPUSERFUNCTIONS_t) As Long
    
  Dim retcode As Long
    
  On Error Resume Next '-- Nothing Will Go Wrong :-)
    
  retcode = 0
  gnZippedFiles = 0
    
  '-- Set Address Of ZIP32.DLL Callback Functions
  '-- (WARNING!) Do Not Change!!!
'  gZipUserFnc.ZDLLPrnt = FnPtr(AddressOf ZDLLPrnt)
'  gZipUserFnc.ZDLLPASSWORD = FnPtr(AddressOf ZDLLPass)
'  gZipUserFnc.ZDLLCOMMENT = FnPtr(AddressOf ZDLLComm)
'  gZipUserFnc.ZDLLSERVICE = FnPtr(AddressOf ZDLLServ)

   If UserFnc.ZDLLPrnt = 0 Then
      UserFnc.ZDLLPrnt = FnPtr(AddressOf ZDLLPrnt)
   End If
   
   If UserFnc.ZDLLPASSWORD = 0 Then
      UserFnc.ZDLLPASSWORD = FnPtr(AddressOf ZDLLPass)
   End If
   
   If UserFnc.ZDLLCOMMENT = 0 Then
      UserFnc.ZDLLCOMMENT = FnPtr(AddressOf ZDLLComm)
   End If
   
   If UserFnc.ZDLLSERVICE = 0 Then
      UserFnc.ZDLLSERVICE = FnPtr(AddressOf ZDLLServ)
   End If
   
   If Err Then
      VBZip32 = Err
      Exit Function
   End If

   DoEvents
   
   '-- Set ZIP32.DLL Callbacks
   retcode = ZpInit(UserFnc)
   
   If Err Then
      VBZip32 = Err
      Exit Function
   End If
      
   If retcode = False Then
      VBZip32 = -1
      Exit Function
   End If
   
   '-- Set ZIP32.DLL Options
   retcode = ZpSetOptions(Opt)
   If retcode = False Then
      VBZip32 = -2
      Exit Function
   End If
     
   '-- Go Zip It Them Up!
   retcode = ZpArchive(nFileNames, ZipFName, FileNames)
   If retcode Then
      VBZip32 = retcode
      Exit Function
   End If
   
   '-- Return The Function Code
   VBZip32 = retcode

End Function


'-- Main ZIP32.DLL Subroutine.
'-- This Is Where It All Happens!!!
'--
'-- (WARNING!) Do Not Change This Function!!!
'--
Public Function VBZip32_old() As Long
    
  Dim retcode As Long
    
  On Error Resume Next '-- Nothing Will Go Wrong :-)
    
  retcode = 0
    
  '-- Set Address Of ZIP32.DLL Callback Functions
  '-- (WARNING!) Do Not Change!!!
'  gZipUserFnc.ZDLLPrnt = FnPtr(AddressOf ZDLLPrnt)
'  gZipUserFnc.ZDLLPASSWORD = FnPtr(AddressOf ZDLLPass)
'  gZipUserFnc.ZDLLCOMMENT = FnPtr(AddressOf ZDLLComm)
'  gZipUserFnc.ZDLLSERVICE = FnPtr(AddressOf ZDLLServ)
    
  '-- Set ZIP32.DLL Callbacks
'  retcode = ZpInit(gZipUserFnc)
    
  '-- Setup ZIP32 Options
  '-- (WARNING!) Do Not Change!
'  ZOPT.Date = zDate                  ' "12/31/79"? US Date?
'  ZOPT.szRootDir = zRootDir          ' Root Directory Pathname
'  ZOPT.szTempDir = zTempDir          ' Temp Directory Pathname
'  ZOPT.fSuffix = zSuffix             ' Include Suffixes (Not Yet Implemented)
'  ZOPT.fEncrypt = zEncrypt           ' 1 If Encryption Wanted
'  ZOPT.fSystem = zSystem             ' 1 To Include System/Hidden Files
'  ZOPT.fVolume = zVolume             ' 1 If Storing Volume Label
'  ZOPT.fExtra = zExtra               ' 1 If Including Extra Attributes
'  ZOPT.fNoDirEntries = zNoDirEntries ' 1 If Ignoring Directory Entries
'  ZOPT.fExcludeDate = zExcludeDate   ' 1 If Excluding Files Earlier Than A Specified Date
'  ZOPT.fIncludeDate = zIncludeDate   ' 1 If Including Files Earlier Than A Specified Date
'  ZOPT.fVerbose = zVerbose           ' 1 If Full Messages Wanted
'  ZOPT.fQuiet = zQuiet               ' 1 If Minimum Messages Wanted
'  ZOPT.fCRLF_LF = zCRLF_LF           ' 1 If Translate CR/LF To LF
'  ZOPT.fLF_CRLF = zLF_CRLF           ' 1 If Translate LF To CR/LF
'  ZOPT.fJunkDir = zJunkDir           ' 1 If Junking Directory Names
'  ZOPT.fGrow = zGrow                 ' 1 If Allow Appending To Zip File
'  ZOPT.fForce = zForce               ' 1 If Making Entries Using DOS Names
'  ZOPT.fMove = zMove                 ' 1 If Deleting Files Added Or Updated
'  ZOPT.fDeleteEntries = zDelEntries  ' 1 If Files Passed Have To Be Deleted
'  ZOPT.fUpdate = zUpdate             ' 1 If Updating Zip File-Overwrite Only If Newer
'  ZOPT.fFreshen = zFreshen           ' 1 If Freshening Zip File-Overwrite Only
'  ZOPT.fJunkSFX = zJunkSFX           ' 1 If Junking SFX Prefix
'  ZOPT.fLatestTime = zLatestTime     ' 1 If Setting Zip File Time To Time Of Latest File In Archive
'  ZOPT.fComment = zComment           ' 1 If Putting Comment In Zip File
'  ZOPT.fOffsets = zOffsets           ' 1 If Updating Archive Offsets For SFX Files
'  ZOPT.fPrivilege = zPrivilege       ' 1 If Not Saving Privelages
'  ZOPT.fEncryption = zEncryption     ' Read Only Property!
'  ZOPT.fRecurse = zRecurse           ' 1 or 2 If Recursing Into Subdirectories
'  ZOPT.fRepair = zRepair             ' 1 = Fix Archive, 2 = Try Harder To Fix
'  ZOPT.flevel = zLevel               ' Compression Level - (0 To 9) Should Be 0!!!
'
  '-- Set ZIP32.DLL Options
  'retcode = ZpSetOptions(gZipOpt)
    
  '-- Go Zip It Them Up!
  retcode = ZpArchive(zArgc, zZipFileName, zZipFileNames)
  
  '-- Return The Function Code
  VBZip32_old = retcode

End Function


