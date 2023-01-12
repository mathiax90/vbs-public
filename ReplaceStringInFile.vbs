Set FSO = CreateObject("Scripting.FileSystemObject")

inFilePath = "C:\test\1.txt"
outFilePath = "C:\test\2.txt"
oldDbPath = "D:\Program Files\Microsoft SQL Server\MSSQL15.MSSQLSERVER\MSSQL\DATA"
newDbPath = "C:\SQLBasa\Data"

Set inFileHandler = FSO.OpenTextFile(inFilePath, 1, false, -1)
Set outFileHandler = FSO.OpenTextFile(outFilePath, 2, true, -1)
j = 0
Do Until inFileHandler.AtEndOfStream	
	inLine = inFileHandler.ReadLine	
	if instr(inLine,oldDbPath) > 0 then
		inLine = replace(inLine, oldDbPath, newDbPath)
		outFileHandler.WriteLine inLine
		j = j + 1
		if j >= 2 then exit do
	else
		outFileHandler.WriteLine inLine
	end if	
Loop

Do Until inFileHandler.AtEndOfStream
	outFileHandler.WriteLine inFileHandler.ReadLine	
Loop


inFileHandler.Close
outFileHandler.Close
msgbox "done!"