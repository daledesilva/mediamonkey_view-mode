' MediaMonkey Script

' NAME: ViewMode
' Version: 1
' Author: Dale de Silva
' Website: www.oiltinman.com
' Date last edited: 20/04/2008

' INSTALL: Copy to Scripts\Auto\

' FILES THAT SHOULD BE PRESENT UPON A FRESH INSTALL:
' ViewMode.vbs

' Thanks to:
' trixmoto (http://trixmoto.net)
' coding the right click menu was only possible through using his RandomisePlaylist 1.1 script as an example

Option Explicit

'Global Variables


Sub firstRun()
	
	SDB.IniFile.BoolValue("ViewMode_MM"&Round(SDB.VersionHi),"AutoLibraryViewModes") = True
		
	OnStartup()
	
End Sub
 
Sub OnStartup()
  'Right Click Node Menu
	createMenu()

	Call Script.RegisterEvent(SDB,"OnChangedSelection","viewModeAdjust")
End Sub


Sub createMenu
	Dim aMnu : Set aMnu = SDB.Objects("ViewMode-RightClickMenu1") 
	Dim aOpt : Set aOpt = SDB.Objects("ViewMode-RightClickOption") 
	
	Set aMnu = SDB.UI.AddMenuItemSub(SDB.UI.Menu_Pop_Tree,0,-1)
	aMnu.Caption = "Default View Mode"
	'itm.OnClickFunc = "ItmClick"
	'itm.UseScript = Script.ScriptPath
	'itm.IconIndex = 25 
	aMnu.Visible = True
	Set SDB.Objects("ViewMode-RightClickMenu1") = aMnu
	
	Set aOpt = SDB.UI.AddMenuItem(aMnu,1,0)
	aOpt.Caption = "Details"
	aOpt.OnClickFunc = "detailsClick"
	aOpt.UseScript = Script.ScriptPath
	Set SDB.Objects("ViewMode-Moption0") = aOpt
	
	Set aOpt = SDB.UI.AddMenuItem(aMnu,1,0)
	aOpt.Caption = "Album Art"
	aOpt.OnClickFunc = "artClick"
	aOpt.UseScript = Script.ScriptPath
	Set SDB.Objects("ViewMode-Moption2") = aOpt
	
	Set aOpt = SDB.UI.AddMenuItem(aMnu,1,0)
	aOpt.Caption = "Album Art with Details"
	aOpt.OnClickFunc = "artDetailsClick"
	aOpt.UseScript = Script.ScriptPath
	Set SDB.Objects("ViewMode-Moption1") = aOpt
	
	Set aOpt = SDB.UI.AddMenuItemSep(aMnu,1,0)
	
	Set aOpt = SDB.UI.AddMenuItem(aMnu,1,0)
	aOpt.Caption = "Clear Default"
	aOpt.OnClickFunc = "modeClearClick"
	aOpt.UseScript = Script.ScriptPath
	
	
	Set aMnu = SDB.UI.AddMenuItemSub(SDB.UI.Menu_Pop_Tree,0,-1)
	aMnu.Caption = "Default Track Browser View"
	Set SDB.Objects("ViewMode-RightClickMenu2") = aMnu
	
	Set aOpt = SDB.UI.AddMenuItem(aMnu,1,0)
	aOpt.Caption = "Visible"
	aOpt.OnClickFunc = "visibleClick"
	aOpt.UseScript = Script.ScriptPath
	Set SDB.Objects("ViewMode-TBoption1") = aOpt
	
	Set aOpt = SDB.UI.AddMenuItem(aMnu,1,0)
	aOpt.Caption = "Hidden"
	aOpt.OnClickFunc = "hiddenClick"
	aOpt.UseScript = Script.ScriptPath
	Set SDB.Objects("ViewMode-TBoption0") = aOpt
	
	Set aOpt = SDB.UI.AddMenuItemSep(aMnu,1,0)
	
	Set aOpt = SDB.UI.AddMenuItem(aMnu,1,0)
	aOpt.Caption = "Clear Default"
	aOpt.OnClickFunc = "browserClearClick"
	aOpt.UseScript = Script.ScriptPath

End Sub



Sub viewModeAdjust
	Dim theVal, theSongList, numAlbums, numArtists
	
	' make sure it doesn't set defaults twice (as when it changes view mode, a changedselection method is called)
	Dim prevNodeCap : prevNodeCap = SDB.IniFile.StringValue("ViewMode_MM"&Round(SDB.VersionHi),"prevNodeCap")
	Dim curNodeCap : curNodeCap = nodePath()
	If prevNodeCap = curNodeCap Then
		Exit Sub
	End If
	SDB.IniFile.StringValue("ViewMode_MM"&Round(SDB.VersionHi),"prevNodeCap") = curNodeCap
	
	'Clear the right click menu default checks
	clearM
	clearTB
	
	' before changing view modes.. check if the current node was using the global default and if it was, remember the global default
	If SDB.IniFile.BoolValue("ViewMode_MM"&Round(SDB.VersionHi),"curMisGlobalDefault") = True Then
		SDB.IniFile.IntValue("ViewMode_MM"&Round(SDB.VersionHi),"GlobalDefaultM") = SDB.MainTracksWindow.ViewMode
	End If
	If SDB.IniFile.BoolValue("ViewMode_MM"&Round(SDB.VersionHi),"curTBisGlobalDefault") = True Then
		SDB.IniFile.IntValue("ViewMode_MM"&Round(SDB.VersionHi),"GlobalDefaultTB") = SDB.MainTracksWindow.TrackBrowserVisibled
	End If
	
	If SDB.IniFile.ValueExists("ViewMode-Nodes_MM"&Round(SDB.VersionHi),nodePath()&"--M") Then
		theVal = SDB.IniFile.IntValue("ViewMode-Nodes_MM"&Round(SDB.VersionHi),nodePath()&"--M")
		changeM(theVal)
		SDB.IniFile.BoolValue("ViewMode_MM"&Round(SDB.VersionHi),"curMisGlobalDefault") = False
	Else
	'---------------------------------------------------------------------------------------------
	' This has been removed because the Albums and Artists Counts are ;
	'		1. reporting the numbers from the PREVIOUS AllVisibleSongList (not the current)
	' and	2. commongly reporting 2 artists when there is only 1
	'---------------------------------------------------------------------------------------------
	
'		If rootNode() = "Library" Then
'			Set theSongList = SDB.AllVisibleSongList
'			numAlbums = theSongList.Albums.Count
'			numArtists = theSongList.Artists.Count
'			MsgBox("numArtists "&numArtists)
'			MsgBox("numAlbums "&numAlbums)
'			If numArtists <= 1 Then
'				If numAlbums <= 1 Then
'					SDB.MainTracksWindow.ViewMode = 0
'				Else
'					SDB.MainTracksWindow.ViewMode = 1
'				End If
'				SDB.IniFile.BoolValue("ViewMode_MM"&Round(SDB.VersionHi),"curMisGlobalDefault") = False
'			Else
'				SDB.MainTracksWindow.ViewMode = SDB.IniFile.IntValue("ViewMode_MM"&Round(SDB.VersionHi),"GlobalDefaultM")
'				SDB.IniFile.BoolValue("ViewMode_MM"&Round(SDB.VersionHi),"curMisGlobalDefault") = True
'			End If
'		Else		
			SDB.MainTracksWindow.ViewMode = SDB.IniFile.IntValue("ViewMode_MM"&Round(SDB.VersionHi),"GlobalDefaultM")
			SDB.IniFile.BoolValue("ViewMode_MM"&Round(SDB.VersionHi),"curMisGlobalDefault") = True
'		End If
	End If
	
	If SDB.IniFile.ValueExists("ViewMode-Nodes_MM"&Round(SDB.VersionHi),nodePath()&"--TB") Then
		theVal = SDB.IniFile.IntValue("ViewMode-Nodes_MM"&Round(SDB.VersionHi),nodePath()&"--TB")
		changeTB(theVal)
		SDB.IniFile.BoolValue("ViewMode_MM"&Round(SDB.VersionHi),"curTBisGlobalDefault") = False
	Else
		SDB.MainTracksWindow.TrackBrowserVisibled = SDB.IniFile.IntValue("ViewMode_MM"&Round(SDB.VersionHi),"GlobalDefaultTB")
		SDB.IniFile.BoolValue("ViewMode_MM"&Round(SDB.VersionHi),"curTBisGlobalDefault") = True
	End If

	showHideMenu()
End Sub


Sub changeM(theVal)
	SDB.MainTracksWindow.ViewMode = theVal
	recordM nodePath(),theVal
	SDB.Objects("ViewMode-Moption"&theVal).Checked = True
End Sub

Sub changeTB(theVal)
	SDB.MainTracksWindow.TrackBrowserVisibled = theVal
	recordTB nodePath(),theVal
	SDB.Objects("ViewMode-TBoption"&theVal).Checked = True
End Sub

Sub clearM
	SDB.Objects("ViewMode-Moption0").Checked = False
	SDB.Objects("ViewMode-Moption1").Checked = False
	SDB.Objects("ViewMode-Moption2").Checked = False
End Sub

Sub clearTB
	SDB.Objects("ViewMode-TBoption0").Checked = False
	SDB.Objects("ViewMode-TBoption1").Checked = False
End Sub

Sub detailsClick(i)
	clearM
	changeM(0)
End Sub


Sub artClick(i)
	clearM
	changeM(2)
End Sub


Sub artDetailsClick(i)
	clearM
	changeM(1)
End Sub


Sub modeClearClick(i)
	clearM
	recordM nodePath(),"clear"
End Sub


Sub visibleClick(i)
	clearTB
	changeTB(1)
End Sub


Sub hiddenClick(i)
	clearTB
	changeTB(0)
End Sub


Sub browserClearClick(i)
	clearTB
	recordTB nodePath(),"clear"
End Sub



Sub recordM(path,theVal)
	path = path & "--M"
	If theVal = "clear" Then
		SDB.IniFile.DeleteKey "ViewMode-Nodes_MM"&Round(SDB.VersionHi),path
		clearM
	Else
		SDB.IniFile.IntValue("ViewMode-Nodes_MM"&Round(SDB.VersionHi),path) = theVal
		SDB.Objects("ViewMode-Moption"&theVal).Checked = True
	End If
End Sub


Sub recordTB(path,theVal)
	path = path & "--TB"
	If theVal = "clear" Then
		SDB.IniFile.DeleteKey "ViewMode-Nodes_MM"&Round(SDB.VersionHi),path
		clearTB
	Else
		SDB.IniFile.IntValue("ViewMode-Nodes_MM"&Round(SDB.VersionHi),path) = theVal
		SDB.Objects("ViewMode-TBoption"&theVal).Checked = True
	End If
End Sub



Sub showHideMenu
	Dim vis : vis = False
	Dim theNode : Set theNode = SDB.MainTree.CurrentNode
	Dim Mnu1 : Set Mnu1 = SDB.Objects("ViewMode-RightClickMenu1") 
	Dim Mnu2 : Set Mnu2 = SDB.Objects("ViewMode-RightClickMenu2")
	
	If Not (theNode Is Nothing) Then
	  'If theNode.NodeType = 61 Or theNode.NodeType = 71 Or theNode.NodeType = 255 Then
	  If theNode.NodeType <> 12 And theNode.NodeType <> 7 And theNode.NodeType <> 6 Then
	  	'playlist, autoplaylist, script created node
	    vis = True
	  End If
	End If
	
	Mnu1.Visible = vis
	Mnu2.Visible = vis
End Sub



Function nodePath
	Dim Tree : Set Tree = SDB.MainTree
	Dim theNode : Set theNode = Tree.CurrentNode
	Dim nodeString
	
	If theNode Is Nothing Then
		nodePath = ""
		Exit Function
	End If
	
	nodeString = theNode.Caption
	Set theNode = Tree.ParentNode(theNode)
	
	Do While theNode.Caption <> ""
		nodeString = theNode.Caption & "-" & nodeString
		Set theNode = Tree.ParentNode(theNode)
	Loop
	nodePath = nodeString
End Function


Function rootNode
	Dim Tree : Set Tree = SDB.MainTree
	Dim theNode : Set theNode = Tree.CurrentNode
	Dim nodeString
	
	nodeString = theNode.Caption
	Set theNode = Tree.ParentNode(theNode)
	
	Do While theNode.Caption <> ""
		nodeString = theNode.Caption
		Set theNode = Tree.ParentNode(theNode)
	Loop
	rootNode = nodeString
End Function
