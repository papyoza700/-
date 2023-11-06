'MoveFolder
Call main

sub main()
  Dim oFSO, sourceFolder, targetFolde
  sourceFolder = "C:\Users\Owner\Desktop\temp1\folder"
  targetFolde = "C:\Users\Owner\Desktop\temp2\folder"
  
  Set oFSO = createobject("Scripting.Filesystemobject")
  oFSO.MoveFolder sourceFolder, targetFolde
  Set oFSO = Nothing
  
  MsgBox "Done",0,"Alert:"
end sub
