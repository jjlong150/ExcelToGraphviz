on runDot(dotParameters)
	set dot_output to do shell script "/usr/local/bin/dot " & dotParameters & " 2>&1"
	return dot_output
end runDot


on chooseOneFile(fileType)
	set resultFile to choose file with prompt "Please choose a " & fileType & " file"
	set resultFile to (the POSIX path of resultFile)
	return resultFile
end chooseOneFile


on chooseImageFile(promptMessage)
	set resultFile to choose file with prompt promptMessage of type {"public.image"}
	set resultFile to (the POSIX path of resultFile)
	return resultFile
end chooseImageFile


on chooseAFolder(currentFolder)
	set theOutputFolder to choose folder with prompt "Please select a folder:"
	set theOutputFolder to (the POSIX path of theOutputFolder)
	return theOutputFolder
end chooseAFolder


on doesFileExist(fileName)
	tell application "System Events"
		if exists file fileName then
			return "true"
		else
			return "false"
		end if
	end tell
end doesFileExist


on doesFolderExist(folderName)
	tell application "System Events"
		if exists folder folderName then
			return "true"
		else
			return "false"
		end if
	end tell
end doesFolderExist


on getSaveAsFileName(extension)
	set resultFile to (choose file name with prompt "Save as " & extension & " File" default name "My File" default location path to desktop) as text
	if resultFile does not end with extension then set resultFile to resultFile & extension
	set resultFile to (the POSIX path of resultFile)
	return resultFile
end getSaveAsFileName


on getVersion(notused)
	return "3.0"
end getVersion


on pickColor(defaultRGB)
	-- Store original delimiters
	set oldDelimiters to AppleScript's text item delimiters
	set AppleScript's text item delimiters to ","
	
	-- Parse input string "R,G,B"
	try
		set rgbList to text items of defaultRGB
		if (count of rgbList) is not equal to 3 then
			set AppleScript's text item delimiters to oldDelimiters
			return ""
		end if
		
		-- Validate and convert to integers
		set r to (item 1 of rgbList) as integer
		set g to (item 2 of rgbList) as integer
		set b to (item 3 of rgbList) as integer
		
		-- Ensure values are in valid range (0-255)
		if r < 0 or r > 255 or g < 0 or g > 255 or b < 0 or b > 255 then
			set AppleScript's text item delimiters to oldDelimiters
			return ""
		end if
		
		-- Convert to macOS 16-bit color space
		set r16 to r * 257
		set g16 to g * 257
		set b16 to b * 257
		
		-- Show color picker
		set pickedColor to choose color default color {r16, g16, b16}
		
		-- Convert back to 8-bit and round to integers
		set rOut to round ((item 1 of pickedColor) / 257)
		set gOut to round ((item 2 of pickedColor) / 257)
		set bOut to round ((item 3 of pickedColor) / 257)
		
		-- Restore delimiters
		set AppleScript's text item delimiters to oldDelimiters
		
		-- Return as comma-separated string
		return (rOut as text) & "," & (gOut as text) & "," & (bOut as text)
		
	on error
		-- Restore delimiters on error
		set AppleScript's text item delimiters to oldDelimiters
		return ""
	end try
end pickColor