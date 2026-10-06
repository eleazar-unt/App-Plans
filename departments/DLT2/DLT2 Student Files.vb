' Written By: Riley Flinn & Andrew Langford
' Version: 2.8 Variant 1 - with Imaging Services changes
' Modified by: Jason Eleazar
' Date Modified: 10/6/2026
'
' This script uses AppGetTree to pull linking information from a browser window and maps it to the document keys.
' this function passes the name of the screen and returns the tree and the tree nodes. treeInfo(0) contains the 
' tree, and treeInfo(1) contains the nodes.
'
'
'--------------------------------------------------------
' INSTRUCTIONS
'--------------------------------------------------------
' 1. Replace "*Host Application*" on line 18 below with the window title of your host application.
' 2. Paste this script in Field1 of the LearnMode Application Plan Designer window and hit okay

' LSPD > Records and Enrollment > Enroll Students > Student Milestones
treeInfo = FindTreeGetNodes("*Student Milestones*")


If treeInfo(0) <> "" Then

  'You should be calling debOuts from our browser script but if you need to call it from here:
  'Call debOut(treeInfo(0), "C:\temp\debout_0.txt")


  '-----------------------CHROME---------------------------
	IF DetermineBrowser() = "chrome" THEN
	
    Field1 = GetValueFromAnchor(treeInfo(1), "[Academic Institution]", -4, 1) 'emplid
    Field2 = GetValueFromAnchor(treeInfo(1), "[Academic Institution]", -7, 1) 'name
    'Field3 = GetValueFromAnchor(treeInfo(1), "[debout anchor]", 1, 1) 
    'Field4 = GetValueFromAnchor(treeInfo(1), "[debout anchor]", 1, 1) 
    'Field5 = GetValueFromAnchor(treeInfo(1), "[debout anchor]", 1, 1)
    'DocProperty("CP NAME HERE") = GetValueFromAnchor(treeInfo(1), "[debout anchor]", 1, 1) 
		
		
  '-----------------------EDGE-----------------------------
  ELSEIF DetermineBrowser() = "edge" THEN

    Field1 = GetValueFromAnchor(treeInfo(1), "[Academic Institution]", -4, 1) 'emplid
    Field2 = GetValueFromAnchor(treeInfo(1), "[Academic Institution]", -7, 1) 'name
    'Field3 = GetValueFromAnchor(treeInfo(1), "[debout anchor]", 1, 1) 
    'Field4 = GetValueFromAnchor(treeInfo(1), "[debout anchor]", 1, 1) 
    'Field5 = GetValueFromAnchor(treeInfo(1), "[debout anchor]", 1, 1)
    'DocProperty("CP NAME HERE") = GetValueFromAnchor(treeInfo(1), "[debout anchor]", 1, 1) 
		
    
  '-----------------------FIRE FOX--------------------------
  ELSEIF DetermineBrowser() = "firefox" THEN

    Field1 = GetValueFromAnchor(treeInfo(1), "[Academic Institution]", -7, 1)
    Field2 = GetValueFromAnchor(treeInfo(1), "[Academic Institution]", -10, 1)
    'Field3 = GetValueFromAnchor(treeInfo(1), "[debout anchor]", 1, 1) 
    'Field4 = GetValueFromAnchor(treeInfo(1), "[debout anchor]", 1, 1) 
    'Field5 = GetValueFromAnchor(treeInfo(1), "[debout anchor]", 1, 1)
    'DocProperty("CP NAME HERE") = GetValueFromAnchor(treeInfo(1), "[debout anchor]", 1, 1) 
	
  
  End If



End If

'--------------------------------------------------------
'IMAGING TEAM FUNCTIONS
'--------------------------------------------------------
'
' Some items in EIS we need are not just in an anchor position.  They are sometimes part
' of the node value.  This gets the value of the node within the tree. 
'
' @Return {String}
function GetValueFromNode(nodes, anchor, position)
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
          GetValueFromNode = GetNodeValue(nodes(anchor_position + position))
     End If
end function

' Called by GetValueFromNode to get the value within the "[]"
'
' @Return {String}
function GetNodeValue(text)
	GetNodeValue = Mid(text, instr(text, "[") + 1, instr(text, "]") - instr(text, "[") - 1)
end function

' This uses the values from the System anchor to determine which browser is being used
'
' @Return {String} firefox | edge | chrome (default)
Function DetermineBrowser()

	browserData = GetValueFromAnchor(treeInfo(1), "[System]", 4, 0)

	Select Case True
		Case InStr(1, browserData, "firefox", vbTextCompare) > 0
			browser = "firefox"
		Case InStr(1, browserData, "edge", vbTextCompare) > 0
			browser = "edge"
		Case Else
			browser = "chrome" 'Chrome is the fallback
	End Select

	DetermineBrowser = browser

End Function

' This is the original GetValueFromAnchor() function from older versions of the script
' which do not contain a directional system.  So you don't have to say up or down along
' with the number of skips.  This is frankly easier for end users to use since it makes
' better logical sense which is why we still use it. 
function GetValueFromAnchorOrig(nodes, anchor, position)
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
          GetValueFromAnchorOrig = GetValue(nodes(anchor_position + position))
     End If
end function


'--------------------------------------------------------
' HYLAND Version 2.8 Code 
' DO NOT EDIT FUNCTIONS BELOW
'--------------------------------------------------------
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
