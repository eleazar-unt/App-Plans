' DO NOT EDIT FUNCTIONS BELOW
function GetValueFromAnchor(nodes, anchor, position, directon)
	anchor_position = -1
	cursor = 0
	
	' searches forwards (starts at lower bound)
	If direction = 0 Then
		for ln = LBound(nodes) to UBound(nodes)
			If InStr(nodes(ln), anchor) > 0 Then
				' Found our anchor.
				anchor_position = cursor
				Exit For
			End If
			cursor = cursor + 1
		next
	' searches backwards (starts at the upper bound)
	ElseIf direction = 1 Then
		for ln = UBound(nodes) to LBound(nodes)
			If InStr(nodes(ln), anchor) > 0 Then
				' Found our anchor.
				anchor_position = cursor
				Exit For
			End If
			cursor = cursor + 1
			
			step -1
		next
	End if
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
		GetValue = Trim(st(1))
	else
		'This gets the value within the "[]"
		if instr(text, "[") > 0 AND instr(text, "]") > 0 Then
			GetValue = Trim(Mid(text, instr(text, "[") + 1, instr(text, "]") - instr(text, "[") - 1))
		else
			GetValue = Trim(text)
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
		' If the above command does not include all the window elements, you can include the "invisible" parameter to get more data from the window.
		'treeChunks = AppGetTree(h, "invisible", treeNodes)
		attemptEnd = attemptEnd + 1
	Wend
	If treeChunks > 0 Then
		for c = 0 to treeChunks - 1
			tree = tree + treeNodes(c)
		next
	Else
		MsgBox Replace(screenName, "*", "") + " browser window was not found after ten attempts!",0,"Notice"
		tree = ""
	End If
	
	'this is for removing elements that are causing issues with linking.
	'tree = replace(tree, "<tag you want to remove"+ vbCrlf,"")
	
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
		'this loop will attempt to bring back a tree up to 45 times, but will stop if it cannot
		for i = 0 to 45
			tree = FindTree(screenName)
			if i = 45 then
				MsgBox "Your linking window may not have loaded. Please attempt again, or contact your ImageNow administrator.",0,"Notice"
				tree = ""
				exit for
			elseif tree = "" then
				exit for
			elseif length > 60 then
				exit for
			else
				' Get nodes within the tree.
				nodes = Split(tree, vbCrLf)
				length = ubound(nodes) + 1
			end if
		next
		FindTreeGetNodes = Array(tree, nodes)
	End If
end function
