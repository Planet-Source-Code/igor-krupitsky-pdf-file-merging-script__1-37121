Set fso = CreateObject("Scripting.FileSystemObject")
sFolder = fso.GetParentFolderName(WScript.ScriptFullName)
Set oFolder = fso.GetFolder(sFolder)
Set oArgs = WScript.Arguments

If oArgs.Count = 0 Then
   'Double Click
   MergeFiles
Else
   'Drag & Drop
   For I = 0 to oArgs.Count - 1
      If LCase(Right(oArgs(I), 4)) = ".pdf" Then
         MergeTwoFiles oArgs(I)
      End If
   Next
End If

'=======================================================
Sub MergeFiles()

    bFirstDoc = True

    If oFolder.Files.Count < 2 Then
        MsgBox "You need to have at least two PDF files in the same folder to merge."
        Exit Sub
    End If

    For Each oFile In oFolder.Files
        If LCase(Right(oFile.Name, 4)) = ".pdf" Then
        
            If bFirstDoc Then
                bFirstDoc = False
                Set oMainDoc = CreateObject("AcroExch.PDDoc")
                oMainDoc.Open sFolder & "\" & oFile.Name
            Else
                Set oTempDoc = CreateObject("AcroExch.PDDoc")
                oTempDoc.Open sFolder & "\" & oFile.Name
                oMainDoc.InsertPages oMainDoc.GetNumPages - 1, oTempDoc, 0, oTempDoc.GetNumPages, False
                oTempDoc.Close
                End If
        End If
    Next
    
    oMainDoc.Save 1, sFolder & "\Output.pdf"
    oMainDoc.Close
    MsgBox "Done! See Output.pdf file."

End Sub
'=======================================================
Sub MergeTwoFiles(sFileName)

   If Not fso.FileExists(sFolder & "\Output.pdf") Then
      fso.CopyFile sFileName, sFolder & "\Output.pdf"
      Exit Sub
   End If

   Set oMainDoc = CreateObject("AcroExch.PDDoc")
   oMainDoc.Open sFolder & "\Output.pdf"

   Set oTempDoc = CreateObject("AcroExch.PDDoc")
   oTempDoc.Open sFileName

   oMainDoc.InsertPages oMainDoc.GetNumPages - 1, oTempDoc, 0, oTempDoc.GetNumPages, False
   oMainDoc.Save 1, sFolder & "\Output.pdf"
   oTempDoc.Close
   oMainDoc.Close
   MsgBox "Done! See Output.pdf file."
End Sub