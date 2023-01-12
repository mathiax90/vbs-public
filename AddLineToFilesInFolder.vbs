Set FSO = CreateObject("Scripting.FileSystemObject")
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8


inDirPath = "g:\test\"

For Each File In FSO.GetFolder(inDirPath).Files
	If UCase(FSO.GetExtensionName(File.Name)) = "MD" Then
		Set FileHandle = FSO.OpenTextFile(File.path, ForAppending)
		filename = File.Name
		
		pos = instr(filename, ".")
		
		tagname = "#" & left(filename,pos-1)
				
		FileHandle.WriteLine
		FileHandle.WriteLine tagname
		FileHandle.Close
	End if
Next

msgbox "done!"