Set fso = CreateObject("Scripting.FileSystemObject")
Dim sFileSuffix: sFileSuffix = "_OCR" 'Appends at the end of output file name
Dim sInFolder: sInFolder = ""
Dim sOutFolder: sOutFolder = ""

if WScript.Arguments.Count <> 1 then
    MsgBox "Please drop folder to move OCR PDF files to " & sFileSuffix & " folder"

Else
    If WScript.Arguments.Count = 1 Then
      sFolder = WScript.Arguments(i)      
      If fso.FolderExists(sFolder) Then  
        sInFolder = sFolder
        sOutFolder = sFolder & sFileSuffix 
        ProcessFolder sFolder        
        MsgBox "Done"
      End If
    End If
End if

Sub ProcessFolder(sFolder)

  iPrefixLen = Len(sFileSuffix) + 4
  sSuffix = Replace(sFolder,sInFolder, "")
  sTargetFolder = sOutFolder & "" & sSuffix

  If fso.FolderExists(sTargetFolder) = False Then
    fso.CreateFolder sTargetFolder
  End If

  Set oFolder = fso.GetFolder(sFolder)
  For Each oFile in oFolder.Files
    
    If Right(oFile.Path, iPrefixLen) = sFileSuffix & ".pdf" Then
      sOutFile = Mid(oFile.Name, 1, Len(oFile.Name) - iPrefixLen) & ".pdf"
      fso.MoveFile oFile.Path, sTargetFolder & "\" & sOutFile
    End If
  Next
  
  For Each oSubfolder in oFolder.SubFolders
    ProcessFolder oSubfolder.Path
  Next
End Sub

