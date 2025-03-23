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
	return "2.0"
end getVersion

