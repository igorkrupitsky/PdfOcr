Set fso = CreateObject("Scripting.FileSystemObject")
Set oShell = WScript.CreateObject ("WSCript.shell")
Dim sFileSuffix: sFileSuffix = "_OCR" 'Appends at the end of output file name
Dim iCount: iCount = 0
Dim oLog
Dim sFolderPath: sFolderPath = GetFolderPath()
sGhostscriptPath = sFolderPath & "\Ghostscript\bin\gswin64.exe" 'https://ghostscript.com/download/gsdnld.html  C:\Program Files\gs\gs10.01.1\bin
sTesseractPath = sFolderPath & "\Tesseract\tesseract.exe"	'https://github.com/UB-Mannheim/tesseract/wiki - C:\Program Files\Tesseract-OCR

sGhostscriptPath = "C:\Program Files\gs\gs10.01.1\bin\gswin64.exe"
sTesseractPath = "C:\Program Files\Tesseract-OCR\tesseract.exe"

if WScript.Arguments.Count = 0 then
    MsgBox "Please drop PDF/image files or folders to convert them to searchable PDFs"

Else
    Set oLog = fso.CreateTextFile(WScript.ScriptFullName & ".log", True)

    For i = 0 to WScript.Arguments.Count -1
      sFile = WScript.Arguments(i)
      
      If fso.FileExists(sFile) Then        
        ProcessFile sFile
        
      ElseIf fso.FolderExists(sFile) Then  
        ProcessFolder sFile
      End If
      
    Next

    oLog.Close    
    MsgBox "Created " & iCount & " searchable PDFs" 
End if

Sub ProcessFolder(sFolder)
  Set oFolder = fso.GetFolder(sFolder)
  For Each oFile in oFolder.Files
    ProcessFile oFile.Path
  Next
  
   For Each oSubfolder in oFolder.SubFolders
    ProcessFolder oSubfolder.Path
   Next
End Sub

Sub ProcessFile(sFile)
    Select Case LCASE(fso.GetExtensionName(sFile))
      Case "pdf"
        OcrPdfFile(sFile)
      Case "bmp","pnm","png","jfif","jpeg","jpg","tiff","gif"
        OcrImgFile(sFile)
    End Select  
End Sub

Sub OcrImgFile(sInFile)
  iPos = InStrRev(sInFile,".")
  sFileBase = Mid(sInFile,1,iPos - 1)
  sOutPdf = sFileBase & sFileSuffix
  
  If fso.FileExists(sOutPdf & ".pdf") Then    
    Msg sOutPdf & ".pdf already exisits"
    Exit Sub
  End If

  oShell.run """" & sTesseractPath & """ """ & sInFile & """ """ & sOutPdf & """ pdf", 1 , True

  If fso.FileExists(sOutPdf & ".pdf") Then
    iCount = iCount + 1
  Else
    Msg sOutPdf & ".pdf could not be created"
  End If  
    
End Sub

Sub OcrPdfFile(sInPdf)
  If Right(sInPdf, Len(sFileSuffix) + 4) = sFileSuffix & ".pdf" Then
      Msg sInPdf & " in an ouput file and will not be processed"
    'Ignore Generated OCR PDFs
    Exit Sub
  End If

  iPos = InStrRev(sInPdf,".")
  sFileBase = Mid(sInPdf,1,iPos - 1)
  sOutPdf = sFileBase & sFileSuffix
  
  If fso.FileExists(sOutPdf & ".pdf") Then    
    Msg sOutPdf & ".pdf already exisits"
    Exit Sub
  End If

  sOutTiff = sFileBase & sFileSuffix & ".tiff"
  oShell.run """" & sGhostscriptPath & """ -dNOPAUSE -q -r300 -sDEVICE=tiff24nc -dBATCH -sOutputFile=""" & sOutTiff & """ """ & sInPdf & """ -c quit", 1 , True
  
  If fso.FileExists(sOutTiff) Then
    oShell.run """" & sTesseractPath & """ """ & sOutTiff & """ """ & sOutPdf & """ pdf", 1 , True
    fso.DeleteFile sOutTiff

    If fso.FileExists(sOutPdf & ".pdf") Then
      iCount = iCount + 1
    Else
      Msg sOutPdf & ".pdf could not be created"
    End If             

  Else
    Msg sOutTiff & " could not be created"
  End If    
End Sub


Function GetFolderPath()
	Dim oFile 'As Scripting.File
	Set oFile = fso.GetFile(WScript.ScriptFullName)
	GetFolderPath = oFile.ParentFolder
End Function

Sub Msg(s)
  oLog.WriteLine Now & vbTab & s
End Sub

