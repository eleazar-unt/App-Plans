' Written By: Riley Flinn & Andrew Langford
' Version: 2.8
' Date Updated: 12/11/2019
' Get values using AppGetTree(). This script uses AppGetTree to pull linking information from a browser window and maps it to the document keys.
' this function passes the name of the screen and returns the tree and the tree nodes. treeInfo(0) contains the tree, and treeInfo(1) contains the nodes.
' *************
' INSTRUCTIONS:
' *************
' 1. Replace "*Host Application*" below with the window title of your host application.
' 2. Paste this script in Field1 of the LearnMode Application Plan Designer window. Click the Test button. That should open Notepad with the accessibility data of the web page. Refer to https://community.hyland.com/gallery/items/60325-how-to-use-the-appgettree-vbscript-function-for-learnmode for full instructions on how to find an anchor in the output and return the value. Example is provided below.
 
treeInfo = FindTreeGetNodes("*[coll18_demo]*")

If treeInfo(0) <> "" Then

	' FOR TESTING ONLY, write the tree to a text file and open it.
	' for generating a tree ensure the following directory exists: "C:\temp\debout_0.txt"
	' if this directory does not exist, then replace the string with a directory that does exist within the quotes.
	'Call debOut(treeInfo(0), "C:\temp\debout_0.txt")

	' Find an anchor label in the output and return the value for that anchor. NOTE: direction to search, direction = 0: forward; direction = 1: backwards)
	'capturedValue = GetValueFromAnchor(treeInfo(1), "YOUR ANCHOR TAG HERE", <+/- number of lines to desired value>,<direction 0 or 1>)

	
	' Example -- the two lines below will search for "Student ID" in the output, get the value that is two lines down, and then assign that to Field1.
	' StudentID = GetValueFromAnchor(treeInfo(1), "Student ID", +2,0)
	' Field1 = StudentID
	
	' To assign a value to a custom property, use the DocProperty array:
	'DocProperty("CP NAME HERE") = variable


	'This assigns values for the SCREEN IDENTIFIER HERE linking Screen
	IF DetermineBrowser() = "chrome" THEN
	
	' potential fix to test is DetermineBrowser().ToLower.Contains("chrome") THEN....
	
		field1 = GetValueFromAnchor(treeInfo(1), "[Item ID]", 3)
		field2 = GetValueFromAnchor(treeInfo(1), "[Name]", 5)
		
		
		' This assigns the values to the Fields
		' Folder = Field1
		' Tab = Field2
		' Field3 = Field3
		' DocProperty("CP NAME HERE") = variable
		Field1 = field1
		Field2 = field2
		
	'This assigns values for the SCREEN IDENTIFIER HERE linking Screen
	ELSE IF DetermineBrowser() = "edge" THEN

		emplid = GetValueFromAnchor(treeInfo(1), "[Empl ID]", 3)
		employeename = GetValueFromAnchor(treeInfo(1), "[Empl ID]", 5)
		
		
		' This assigns the values to the Fields
		' Folder = Field1
		' Tab = Field2
		' Field3 = Field3
		' DocProperty("CP NAME HERE") = variable
		Field1 = emplid
		Field2 = employeename
		Field4 = GradYear
		
	' NOTE: Firefox if picky about the "dash" that is used.  make sure you are using the "long"
    '       version of the dash as seen below.  " - " is for Chrome/Edge, " — " is for Firefox. 
	ELSE IF DetermineBrowser() = "firefox" THEN

		' this uses "YOUR ANCHOR TAG HERE" as the anchor and indexes the capturedValue off this anchor tag.
		emplid = GetValueFromAnchor(treeInfo(1), "[Empl ID]", 4)
		employeename = GetValueFromAnchor(treeInfo(1), "[Empl ID]", 9)
		
				' This assigns the values to the Fields
		' Folder = Field1
		' Tab = Field2
		' Field3 = Field3
		' DocProperty("CP NAME HERE") = variable
		Field1 = emplid
		Field2 = employeename
		
		
	' NOTE: For further linking screens, please add more if statements based on the ones listed above.
	End If
	End If
	End If


End If


'IMAGING TEAM FUNCTIONS
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

' Function get the value of the node within the tree.
function GetNodeValue(text)

    'This gets the value within the "[]"
	GetNodeValue = Mid(text, instr(text, "[") + 1, instr(text, "]") - instr(text, "[") - 1)
	
end function

' This uses the values from the System anchor to determine which browser is being used
' @Return {String} firefox | edge | chrome (default)
Function DetermineBrowser()

	browserData = GetValueFromAnchor(treeInfo(1), "[System]", 8)

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
