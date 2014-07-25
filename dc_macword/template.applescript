tell application "Microsoft Word"
	-- retrieving the path of the current document in word
	set docPath to the full name of the active document
	-- converting it to POSIX form, in order to manipulate it
	set docPath to POSIX path of docPath

	-- retrieving the doc title
	set docName to the name of the active document

end tell

-- defining a temporary folder
set tempDir to path to temporary items for user domain
-- converting the tempDir path to POSIX form, in order to manipulate it
set tempDir to POSIX path of tempDir

-- generating a path to a temporary file
set tempFile to quoted form of tempDir & "~" & docName

-- retrieving path of doccleaner
set py to "doccleaner.py "
workingDir = $DOCCLEANER_PATH

set callDir to quoted form of workingDir & py

-- launching doccleaner
do shell script callDir -i docPath -o tempFile -t $XSL_PATH

-- TODO : copying content from tempFile to original doc
tell application "Microsoft Word"

end tell

-- removing the tempFile
tell application "Finder"
	delete file tempFile
	-- empty trash ?
end tell