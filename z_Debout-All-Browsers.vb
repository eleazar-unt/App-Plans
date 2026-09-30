' Written By: Riley Flinn
' Version: 2.4
' Date Updated: 11/07/2018
' Last Updated by: Riley Flinn & Andrew Langford
' Modified by:
' Date Modified:

' Get values using AppGetTree(). This script uses AppGetTree to pull linking information from a browser window and maps it to the document keys.

' this function passes the name of the screen and returns the tree and the tree nodes. treeInfo(0) contains the tree, and treeInfo(1) contains the nodes.
' Replace "*Host Application*" with the window title of your host application
treeInfo = FindTreeGetNodes("*Modify a Person*")

	
	
If treeInfo(0) <> "" Then

	browser = GetValueFromAnchor(treeInfo(1), "[System]", 2)
	FirefoxMatch= "Mozilla Firefox"
	EdgeMatch="Edge"
	ChromeMatch="Google Chrome"
        
	FirefoxFound = InStr(browser, FirefoxMatch)
	EdgeFound = InStr(browser, EdgeMatch)
	ChromeFound = InStr(browser, ChromeMatch)
	
	If (FirefoxFound > 0 ) Then
		call debOut(treeInfo(0), "C:\temp\debout_Firefox.txt")
	ElseIf (EdgeFound > 0) Then
		call debOut(treeInfo(0), "C:\temp\debout_Edge.txt")
	ElseIf (ChromeFound > 0) Then
		call debOut(treeInfo(0), "C:\temp\debout_Chrome.txt")
	End If

End If


' DO NOT EDIT FUNCTIONS BELOW
function GetValueFromAnchor(nodes, anchor, position)
     anchor_position = -1
     cursor = 0
     for ln = LBound(nodes) to UBound(nodes)
          If InStr(nodes(ln), anchor) > 0 Then
               ' Found our anchor.
               anchor_position = cursor
               Exit For
          End If
          cursor = cursor + 1
     next
     If anchor_position = -1 Then
          MsgBox("Could not find " & anchor & ".")
     Else
          GetValueFromAnchor = GetValue(nodes(anchor_position + position))
     End If
end function

' Function to write tree to a text file
function debOut(tree, tempDirectory)

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    fcount = 0
    exists = objFSO.FileExists(tempDirectory)
    while exists and fcount < 10
        fcount = fcount + 1
        s = "_" + CStr(fcount)
        tempDirectory = Replace(tempDirectory, "_" + CStr(fcount - 1), "_" + CStr(fcount))
        exists = objFSO.FileExists(tempDirectory)
    wend
    Set objFile = objFSO.CreateTextFile(tempDirectory, True, true)
    objFile.Write(tree)
    objFile.Close
    CreateObject("WScript.Shell").Run(tempDirectory)
	
end function

' Function get the value of the node within the tree.
function GetValue(text)
    st = Split(text, "->")
    if UBound(st) > 0 then
        GetValue = st(1)
    else
        'This gets the value within the "[]"
        if instr(text, "[") > 0 AND instr(text, "]") > 0 Then
            GetValue = Mid(text, instr(text, "[") + 1, instr(text, "]") - instr(text, "[") - 1)
        else
            GetValue = text
        end if
    end if
end function

'this function returns the tree
function FindTree(screenName)
	'this will loop up to 10 times trying to get the tree.
	attemptEnd = 0
	h = FindWindow(screenName)
	While treeChunks <= 0 OR attemptEnd = 9
		treeChunks = AppGetTree(h, "", treeNodes)
		attemptEnd = attemptEnd + 1
	Wend
	If treeChunks > 0 Then
		for c = 0 to treeChunks - 1
			tree = tree + treeNodes(c)
		next
	Else
		MsgBox Replace(screenName, "*", "") + " Chrome window was not found after ten attempts!",0,"Notice"
		tree = ""
	End If
	
	findTree = tree
end function

'this function handles getting the nodes.
function FindTreeGetNodes(screenName)

	' This section handles if there is a shortened tree that is returned by the FindTree function
	h = FindWindow(screenName)
	If h = 0 Then
		MsgBox "A window with the title, " + Replace(screenName, "*", "") + ", was not found.",0,"Notice"
		FindTreeGetNodes = Array("","")
	Else
	
		While length <= 60
			tree = FindTree(screenName)
			if tree = "" then
				length = 65
			else
				' Get nodes within the tree.
				nodes = Split(tree, vbCrLf)
				length = ubound(nodes) + 1
			end if
		Wend
		
		FindTreeGetNodes = Array(tree, nodes)
		
	End If
end function
