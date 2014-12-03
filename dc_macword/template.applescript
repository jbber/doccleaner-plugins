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
	
	set transitionalDoc to originalDocPath
	--loop here
	set jj to 1
	
	repeat $XSLNUMBER times

		if (jj > 1) then	
			set transitionalDoc to tempFile
			set tempFile to quoted form of tempDir & "~" & jj & originalDocName
			set tempFile to POSIX file tempFile as alias
		end if
		set xslpath to item XSL_PATH of item jj of $PROCESSINGS
		set subfile to item SUBFILE of item jj of $PROCESSINGS
		set xslparameter to item XSLPARAMETER of item jj of $PROCESSINGS
		if $XSLPARAMETER != "0" 
			set doccleaner to quoted form of workingDir & py & " -i " & transitionalDoc & " -o " & tempFile & " -t " & XSL_PATH & " -s " & SUBFILE & " - p "& XSLPARAMETER
		else
			set doccleaner to quoted form of workingDir & py & " -i " & transitionalDoc & " -o " & tempFile & " -t " & XSL_PATH & " -s " & SUBFILE
		end if
		-- launching doccleaner
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
		
		set jj to (jj + 1)
	end repeat
	
end tell

-- removing the tempFile
tell application "Finder"
	delete file tempFile
	-- empty trash ?
end tell