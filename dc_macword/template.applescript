tell application "Microsoft Word"
	-- TODO : test on MacWord 2004 & 2008 

	-- retrieving info about the current document
	set originalDocPath to the full name of the active document
	set originalDocPath to POSIX path of originalDocPath
	set originalDocName to the name of the active document
	
	-- setting variables to use the temporary file
	set tempDir to path to temporary items from user domain
	
	-- converting the tempDir path to POSIX form, in order to manipulate it
	set tempDir to POSIX path of tempDir
	
	-- generating a path to a temporary file
	set tempFile to quoted form of tempDir & "~" & originalDocName
	set tempFile to POSIX file tempFile as alias
	
	-- retrieving path of doccleaner
	set py to "doccleaner.py "
	workingDir = $DOCCLEANER_PATH
	set doccleaner to quoted form of workingDir & py & " -i " & originalDocPath & " -o " & tempFile & " -t " & $XSL_PATH
	
	-- launching doccleaner
	-- TODO: handling the parameters --subfile and --xslparameters
	do shell script doccleaner
	
	-- opening the tempFile created by doccleaner
	activate (open tempFile)
	
	-- passing its content to the newContent variable
	set newContent to formatted text of text object of document tempFile
	
	-- reactivate initial doc
	activate (open originalDocPath)
	
	-- copy content from tempFile to originalDoc, without using the clipboard (it would be bad practice)
	set formatted text of text object of active document to newContent
	
	-- closing the temp file
	close document tempFile
end tell

-- removing the tempFile
tell application "Finder"
	delete file tempFile
	-- empty trash ?
end tell