' This script exports the complete MediaMonkey database (songs and playlist) into
' the iTunes xml format. Some caveats apply, for details and the latest version
' see the MediaMonkey forum thread at
' http://www.mediamonkey.com/forum/viewtopic.php?f=2&t=31680
'
' Change history:
' 1.0   initial version by "DC"
' 1.1   options added for disabling timer and showing a file selection dialog
' 1.2   fixed: unicode characters (e.g. Chinese) were encoded different than iTunes does
' 1.3   fixed: handling of & and # in URI encoding, added Last Played
' 1.4   fixed: Traktor failing import due to invalid characters in xml (& -> &#38;)
' 1.5   added BPM field, added forced export on shutdown (Matthias, 12.12.2012)
'       added child-playlists (Matthias, 12.12.2012)
' 1.6   migrate from report to MediaMonkey plugin with MMIP installer
' 1.6.1 improve unicode utf-8 output; add handling of utf-16 surrogate pairs
' 1.6.2 added Options dialog
'       dynamically configurable options for export at shutdown and periodic export
'       dynamically configurable filename and directory

option explicit     ' report undefined variables, ...

'  ------------------------------------------------------------------
const EXPORTING = "itunes_export_active"
dim scriptControl ' : scriptControl = CreateObject("ScriptControl")

' Returns encoded URI for provided location string. 
function encodeLocation(location)
  ' 10.10.2010: need jscript engine to access its encodeURI function which is not 
  ' available in vbscript
  if isEmpty(scriptControl) then
    set scriptControl = CreateObject("ScriptControl")
    scriptControl.Language = "JScript"
  end if

  location = replace(location, "\", "/")
  location = replace(location, "&", "&")
  encodeLocation = scriptControl.Run("encodeURI", location)
  encodeLocation = replace(encodeLocation, "#", "%23")
end function

' Returns UTF8 equivalent string of the provided Unicode codepoint c.
' For the argument AscW should be used to get the Unicode codepoint
' (not Asc).
' Function by "Arnout", copied from this stackoverflow question:
' http://stackoverflow.com/questions/378850/utf-8-file-appending-in-vbscript-classicasp-can-it-be-done
function Utf8(ByVal c)
  dim b1, b2, b3, b4
  if c < 128 then ' 1 byte utf-8
    Utf8 = chr(c)
  elseif c < 2048 then ' 2 byte utf-8
    b1 = c mod 64
    b2 = (c - b1) / 64
    Utf8 = chr(&hc0 + b2) & chr(&h80 + b1)

  elseif c < 65536 Then ' 3 byte utf-8
    b1 = c mod 64
    b2 = ((c - b1) / 64) mod 64
    b3 = (c - b1 - (64 * b2)) / 4096
    Utf8 = chr(&he0 + b3) & chr(&h80 + b2) & chr(&h80 + b1)

  elseif c < &h10ffff& then ' 4 byte utf-8
    b1 = c mod 64
    b2 = ((c - b1) / 64 ) mod 64
    b3 = ((c - b1 - (64 * b2)) / 4096) mod 64
    b4 = ((c - b1 - (64 * b2) - (4096 * b3)) / 262144)
    Utf8 = chr(&hf0 + b4) & chr(&h80 + b3) & chr(&h80 + b2) & chr(&h80 + b1)
    
  else ' error - use replacement character
    Utf8 = chr(&hef) & chr(&hbf) & chr(&hdb)
  end if
end function

' Returns the XML suitable escaped version of the srcstring parameter.
' This function is based on MapXML found in other MM scripts, e.g.
' Export.vbs, but fixes a unicode issue and is probably faster.
' Note that a bug in AscW still prevents the correct handling of unicode
' codepoints > 65535.
'
' added escaping of xml special characters as per original itunes and required by Traktor parser
function escapeXML(srcstring)
  dim i, codepoint, currentchar, replacement
  i = 1
  while i <= Len(srcstring)
    currentchar = mid(srcstring, i, 1)
    replacement = null
    if currentchar = "&" then
      replacement = "&#38;"
    elseif currentchar = "<" then
      replacement = "&#60;"
    elseif currentchar = ">" then
      replacement = "&#62;"
    elseif currentchar =  CHR(34) then
      replacement = "&#34;"
    else
      codepoint = (AscW(currentchar) And &hffff&)

      ' Handle surrogate pairs; see https://unicodebook.readthedocs.io/unicode_encodings.html#utf-16-surrogate-pairs
      if codepoint >= &hd800& and codepoint <= &hdbff& then
        dim codepoint2 ' for lower-pair
        codepoint2 = (AscW(mid(srcstring, i+1, 1)) And &hffff&)
        codepoint = &h10000& + Clng((codepoint and &h3ff) * 1024) + Clng(codepoint2 and &h3ff)

        ' remove the 2nd code (lower-pair) from the string
        srcstring = mid(srcstring, 1, i)  + Mid(srcstring, i + 2, Len(srcstring))
      end if
      
      ' Important: reject control characters except tab, cr, lf. See also http://www.w3.org/TR/1998/REC-xml-19980210.html#NT-Char
      if codepoint > 127 or currentchar = vbTab or currentchar = vbLf or currentchar = vbCr then
        replacement = Utf8(codepoint)
      elseif codepoint < 32 then
        replacement = ""
      end if
    end if
    
    if not IsNull(replacement) then ' otherwise we keep the original srcstring character (common case)
      srcstring = mid(srcstring, 1, i - 1) + replacement + Mid(srcstring, i + 1, Len(srcstring))
      i = i + len(replacement)
    else
      i = i + 1
    end if
  wend
  escapeXML = srcstring
end function


' Getter for the configured ExportAtShutdown boolean
function getExportAtShutdown()
  dim myIni
  dim myValue
  dim myBool
  
  set myIni = SDB.IniFile
  myValue = cleanFilename(myIni.StringValue("ExportITunesXML","ExportAtShutdown"))
  'MsgBox "DBG getExportAtShutdown(): '" & myValue & "'"

  ' parse ini value to boolean; use default if not defined as 0/1
  if myValue = "0" then 
    myBool = False
  elseif myValue = "1" then
    myBool = True
  else
    myBool = getDefaultExportAtShutdown()
  end if

  'MsgBox "DBG getExportAtShutdown(): '" & myBool & "'"
  getExportAtShutdown = myBool
end function
'
' Setter for the configured ExportAtShutdown boolean
sub setExportAtShutdown(byVal myBool)
  dim myIni
  set myIni = SDB.IniFile
  'MsgBox "DBG setExportAtShutdown(): " & myBool

  if myBool then
    myIni.StringValue("ExportITunesXML","ExportAtShutdown") = "1"
  else
    myIni.StringValue("ExportITunesXML","ExportAtShutdown") = "0"
  end if
end sub
'
function getDefaultExportAtShutdown()
  getDefaultExportAtShutdown = False
end function 

' Getter for the configured PeriodicExport boolean
function getPeriodicExport()
  dim myIni
  dim myValue
  dim myBool
  
  set myIni = SDB.IniFile
  myValue = cleanFilename(myIni.StringValue("ExportITunesXML","PeriodicExport"))
  'MsgBox "DBG getPeriodicExport(): '" & myValue & "'"

  ' parse ini value to boolean; use default if not defined as 0/1
  if myValue = "0" then 
    myBool = False
  elseif myValue = "1" then
    myBool = True
  else
    myBool = getDefaultPeriodicExport() 
  end if

  'MsgBox "DBG getPeriodicExport(): '" & myBool & "'"
  getPeriodicExport = myBool
end function
'
' Setter for the configured PeriodicExport boolean
sub setPeriodicExport(byVal myBool)
  dim myIni
  set myIni = SDB.IniFile
  'MsgBox "DBG getPeriodicExport(): " & myBool

  if myBool then
    myIni.StringValue("ExportITunesXML","PeriodicExport") = "1"
  else
    myIni.StringValue("ExportITunesXML","PeriodicExport") = "0"
  end if
end sub
'
function getDefaultPeriodicExport()
  getDefaultPeriodicExport = False
end function 



' Getter for the configured Directory
function getDirectory()
  ' FIXME - for now, simply return the default....
    'edt.Text = ini.StringValue("ExportITunesXML","Directory")    
  getDirectory = getDefaultDirectory()
end function

' Setter for the configured Directory
sub setDirectory(byVal myDirectory)

  ' FIXME
  ' do so basic cleanup; ensure the path ends with a directory separator
  if right(myDirectory,1) <> "\" then
    ' simply append the missing separator
    myDirectory = myDirectory & "\"
  end if 
MsgBox "DBG setDirectory(): " & myDirectory


  ' only store if valid!
  ' Dim ini : Set ini = SDB.IniFile
  'ini.StringValue("ExportITunesXML","Site") = Sheet.Common.ChildControl("NPSite").Text
  'ini.StringValue("ExportITunesXML","User") = Sheet.Common.ChildControl("NPUser").Text
  'ini.StringValue("ExportITunesXML","Path") = Sheet.Common.ChildControl("NPPath").Text

end sub

' Get default for the configured Directory
function getDefaultDirectory()
  ' The default file location will be in the same folder as the database 
  ' because this folder is writable and user specific.
  dim dbpath : dbpath = SDB.Database.Path
  dim parts : parts = split(dbpath, "\")
  dim dbfilename : dbfilename = parts(UBound(parts))
  dim path : path = Mid(dbpath, 1, Len(dbpath) - Len(dbfilename))
  getDefaultDirectory = path
end function

' Return true if the directory is valid and writable by the user
function isValidDirectory(byVal myDirectory)
  ' FIXME - for now, simply assume it is...
  isValidDirectory = True
end function 


' Getter for the configured Filename
' if filename is undefined/blank then return the default
function getFilename()
  dim myIni
  dim myFilename 
  
  set myIni = SDB.IniFile
  myFilename = cleanFilename(myIni.StringValue("ExportITunesXML","Filename"))
  
  'MsgBox "DBG getFilename(): '" & myFilename & "'"
  if myFilename = "" then 
    myFilename = getDefaultFilename() 
  end if

  getFilename = myFilename
end function

' Setter for the configured Filename
sub setFilename(byVal myFilename)
  dim myIni

  ' trim any unsupported characters:
  myFilename = cleanFilename(myFilename)

  'MsgBox "DBG setFilename(): " & myFilename
  set myIni = SDB.IniFile
  myIni.StringValue("ExportITunesXML","Filename") = myFilename

end sub

' Get default for the configured Filename
function getDefaultFilename()
  ' The default filename will be same as written by Apple iTunes
  getDefaultFilename = "iTunes Music Library.xml"
end function

' remove invalid characters from the filename
function cleanFilename(byVal myFilename)
  Const sInvalidChars = "/\|<>:*?"""
  Dim idx
  for idx = 1 to len(sInvalidChars)
    myFilename = replace(myFilename, mid(sInvalidChars, idx, 1), "")
  next
 cleanFilename = trim(myFilename)
End Function


' N must be numberic. Return value is N converted to a string, padded with
' a single "0" if N has only one digit.
function LdgZ(N)    
  if (N >= 0) and (N < 10) then 
    LdgZ = "0" & N 
  else 
    LdgZ = "" & N  
  end if  
end function  

' Adds a simple key/value pair to the XML accessible via textfile fout.
sub addKey(fout, key, val, keytype)
  if keytype = "string" then
    if val = "" then ' nested if because there is no shortcut boolean eval
      exit sub
    end if
  end if
  
  if keytype = "integer" then
    if val = 0 then ' nested if because there is no shortcut boolean eval
      exit sub
    end if
  end if
  
  if keytype = "date" then ' convert date into ISO-8601 format
    val = Year(val) & "-" & LdgZ(Month(val)) & "-" & LdgZ(Day(val)) _
      & "T" & LdgZ(Hour(val)) &  ":" & LdgZ(Minute(val)) & ":" & LdgZ(Second(val))
  end if
  
  fout.WriteLine "            <key>" & key & "</key><" & keytype & ">" & val & "</" & keytype & ">"
end sub

' Return the full path of the file to export to. The file will be located 
' in the same folder as the database because this folder is writable and user
' specific. For maximum compatibility we will use the original iTunes name
' which is "iTunes Music Library.xml".
' 29.03.2009: if the new option QUERY_FOLDER is set to true this function
' will query for the folder to save to instead.
function getExportFilename()
'  dim path
'  if QUERY_FOLDER then
'    dim inif
'    set inif = SDB.IniFile
'    path = inif.StringValue("Scripts", "LastExportITunesXMLDir")
'    path = SDB.SelectFolder(path, SDB.Localize("Select where to export the iTunes XML file to."))
'    if path = "" then
'      exit function
'    end if
'    if right(path, 1) <> "\" then
'      path = path & "\"
'    end if
'    inif.StringValue("Scripts", "LastExportITunesXMLDir") = path
'    set inif = Nothing  
'  else
'    dim dbpath : dbpath = SDB.Database.Path
'    dim parts : parts = split(dbpath, "\")
'    dim dbfilename : dbfilename = parts(UBound(parts))
'    path = Mid(dbpath, 1, Len(dbpath) - Len(dbfilename))
'  end if
'  getExportFilename = path + "iTunes Music Library.xml"
  getExportFilename = getDirectory() + getFilename()
end function

' MM stores childplaylists, while iTunes XML stores parent playlist
' This function gets the parent playlist (if existent) 
' Added 12.12.2012 by Matthias 
function getparentID(playlist)
	Dim childID, childItems, i, iter
	childID = playlist.ID
    set iter = SDB.Database.OpenSQL("select PlaylistName from PLAYLISTS")
    while not iter.EOF 
		if playlist.Title <> "Accessible Tracks" then ' this would correspond to iTunes' "Library" playlist
			Set childItems = SDB.PlaylistByTitle(iter.StringByIndex(0)).ChildPlaylists
			For i=0 To childItems.Count-1
				if childItems.Item(i).ID = childID then  
					getparentID = SDB.PlaylistByTitle(iter.StringByIndex(0)).ID
					exit function
				end if 
			next
		end if
      iter.next
    wend      
    set iter = nothing
	getparentID = 0 
end function


' Exports the full MM library and playlists into an iTunes compatible
' library.xml. This is not intended to make MM's database available to
' iTunes itself but to provide a bridge to other applications which are
' able to read the iTunes library xml.
sub Export
  if SDB.Objects(EXPORTING) is nothing then
    SDB.Objects(EXPORTING) = SDB
  else
    MsgBox SDB.Localize("iTunes export is already in progress."), 64, "iTunes Export Script"
    exit sub
  end if

  dim filename, fso, iter, songCount, fout, progress, song, playlistCount
  dim progressText, i, j, tracks, playlist
  
  filename = getExportFilename()
  if filename = "" then
    SDB.Objects(EXPORTING) = nothing
    exit sub
  end if

  set fso = SDB.Tools.FileSystem
  set fout = fso.CreateTextFile(filename, true)

  set iter = SDB.Database.OpenSQL("select count(*) from SONGS")
  songCount = Int(iter.ValueByIndex(0)) ' needed for progress
  set iter = SDB.Database.OpenSQL("select count(*) from PLAYLISTS")
  playlistCount = CInt(iter.ValueByIndex(0)) 

  set progress = SDB.Progress
  progressText = SDB.Localize("Exporting to " & getFilename() & "...")
  Progress.Text = progressText
  Progress.MaxValue = songCount + playlistCount * 50

  fout.WriteLine "<?xml version=""1.0"" encoding=""UTF-8""?>"
  fout.WriteLine "<!DOCTYPE plist PUBLIC ""-//Apple Computer//DTD PLIST 1.0//EN"" ""http://www.apple.com/DTDs/PropertyList-1.0.dtd"">"
  fout.WriteLine "<plist version=""1.0"">"
  fout.WriteLine "<dict>"
  fout.WriteLine "    <key>Major Version</key><integer>1</integer>"
  fout.WriteLine "    <key>Minor Version</key><integer>1</integer>"
  fout.WriteLine "    <key>Application Version</key><string>7.6</string>"
  fout.WriteLine "    <key>Features</key><integer>5</integer>" ' whatever that means
  fout.WriteLine "    <key>Show Content Ratings</key><true/>"
  ' Fields not available in MM:
  ' fout.WriteLine "    <key>Music Folder</key><string>file://localhost/C:/....../iTunes/iTunes%20Music/</string>"
  ' fout.WriteLine "    <key>Library Persistent ID</key><string>4A9134D6F642512F</string>"

  ' Songs
  ' 
  ' For each song write available tag values to the library.xml. At this time 
  ' this does not include artwork, volume leveling and album rating.
  if songCount > 0 then
    fout.WriteLine "    <key>Tracks</key>"
    fout.WriteLine "    <dict>"
    i = 0
    set iter = SDB.Database.QuerySongs("")
    while not iter.EOF and not Progress.Terminate and not Script.Terminate
      set song = iter.Item
      iter.next

      ' %d always inserts 0, don't know why
      i = i + 1
      progress.Text = progressText & " " & SDB.LocalizedFormat("%s / %s songs", CStr(i), CStr(songCount), 0)
      if i mod 50 = 0 then
        SDB.ProcessMessages
      end if

      fout.WriteLine "        <key>" & Song.id & "</key>"
      fout.WriteLine "        <dict>   "
      addKey fout, "Track ID", Song.id, "integer"
      addKey fout, "Name", escapeXML(Song.Title), "string"
      addKey fout, "Artist", escapeXML(Song.ArtistName), "string"
      addKey fout, "Composer", escapeXML(Song.MusicComposer), "string"
      addKey fout, "Album Artist", escapeXML(Song.AlbumArtistName), "string"
      addKey fout, "Album", escapeXML(Song.AlbumName), "string"
      addKey fout, "Kind", escapeXML("MPEG audio file"), "string"
      addKey fout, "Size", Song.FileLength, "integer"
      addKey fout, "Genre", escapeXML(Song.Genre), "string"
      addKey fout, "Total Time", Song.SongLength, "integer"
      addKey fout, "Track Number", Song.TrackOrder, "integer" ' potential type problem with TrackOrderStr
      addKey fout, "Disc Number", Song.DiscNumber, "integer" ' potential type problem with DiscNumberStr
      addKey fout, "Play Count", Song.PlayCounter, "integer"
      if Song.Rating >= 0 and Song.Rating <= 100 then
        addKey fout, "Rating", Song.Rating, "integer" ' rating seems to be compatible in range (although not stored in same id3 tag)
      end if
      addKey fout, "Year", Song.Year, "integer"
      addKey fout, "Date Modified", Song.FileModified, "date"
      addKey fout, "Date Added", Song.DateAdded, "date"
      addKey fout, "Play Date UTC", Song.LastPlayed, "date"
      addKey fout, "Bit Rate", Int(Song.Bitrate / 1000), "integer"
      addKey fout, "Sample Rate", Song.SampleRate, "integer"
      addKey fout, "Track Type", escapeXML("File"), "string"
      addKey fout, "File Folder Count", -1, "integer"
      addKey fout, "Library Folder Count", -1, "integer"
      addKey fout, "Comments", escapeXML(Song.Comment), "string"
      addKey fout, "BPM", Song.BPM, "string"
      
      ' 10.10.2010: fixed: location was not correctly URI encoded before
      ' addKey fout, "Location", "file://localhost/" & Replace(Replace(Escape(Song.Path), "%5C", "/"), "%3A", ":"), "string"
      ' addKey fout, "Location", encodeLocation("file://localhost/" & Song.Path), "string"
      ' 04.07.2018: amparsant needs to be escaped
      addKey fout, "Location", Replace(encodeLocation("file://localhost/" & Song.Path),"&","&#38;"), "string"

      ' TODO artwork?
      ' addKey fout, "Artwork Count", 0, "integer"
      ' TODO convert to iTunes rating range. MM: -99999...?. iTunes: -255 (silent) .. 255
      ' fout.WriteLine "            <key>Volume Adjustment</key><integer>" & escapeXML(Song.Leveling) & "</integer>" 

      ' Fields not available in MM:
      ' fout.WriteLine "            <key>Disc Count</key><integer>" & escapeXML(Song.?) & "</integer>"
      ' fout.WriteLine "            <key>Album Rating</key><integer>" & escapeXML(Song.?) & "</integer>"
      ' fout.WriteLine "            <key>Persistent ID</key><string>5282DFDE369975A8</string>"

      fout.WriteLine "        </dict>"

      Progress.Increase
    wend
    fout.WriteLine "    </dict>"
  end if
  SDB.ProcessMessages
  
  ' Playlists
  '
  ' This part differs at least with the following items from an original iTunes 
  ' library.xml:
  ' - iTunes includes a playlist named "Library" with all songs, we don't
  ' - every iTunes playlist has a "Playlist Persistent ID", e.g. "4A9134D6F6425130"
  '   We don't have that data.
  '
  ' Also note: auto-playlists are evaluated once and are exported like that. They
  ' are not converted into iTunes auto-playlists. A consequence of this is that
  ' e.g. randomized or size-limited playlists will contain a static snapshot taken
  ' at export time.
  if playlistCount > 0 and not Progress.Terminate and not Script.Terminate then
    fout.WriteLine "    <key>Playlists</key>"
    fout.WriteLine "    <array>"
    
    ' Get playlists and store them into an array. Make sure that we do not have
    ' an open query while playlist.Tracks is evaluated because that will fail
    ' (it wants to start a db transaction but can't because a query is still open)
    dim playlists()
    set iter = SDB.Database.OpenSQL("select PlaylistName from PLAYLISTS")
    i = 0
    while not iter.EOF 
      set playlist = SDB.PlaylistByTitle(iter.StringByIndex(0))
      if playlist.Title <> "Accessible Tracks" then ' this would correspond to iTunes' "Library" playlist
        redim preserve playlists(i)
        set playlists(i) = playlist
        i = i + 1
      end if
      iter.next
    wend      
    set iter = nothing

    for each playlist in playlists
	    dim parentID
	    parentID = getparentID(playlist)
      set tracks = playlist.Tracks
      ' %d always inserts 0, don't know why
      i = i + 1
      progress.Text = progressText & " " & SDB.LocalizedFormat("playlist ""%s"" (%s songs)", playlist.Title, CStr(tracks.Count), 0)
      SDB.ProcessMessages

      fout.WriteLine "        <dict>"
      addKey fout, "Name", escapeXML(playlist.Title), "string"
      ' Apparently only used for "Library" playlist:
      ' addKey fout, "Master", Nothing, "true"
      ' addKey fout, "Visible", Nothing, "empty"
      addKey fout, "Playlist ID", playlist.ID, "integer"
      ' No MM field for this:
      
	  addKey fout, "Playlist Persistent ID", playlist.ID, "string"
	  if parentID <> 0 then
		  addKey fout, "Parent Persistent ID", parentID, "string"
	  end if 
      fout.WriteLine "         <key>All Items</key><true/>"
      if tracks.Count > 0 then      
        fout.WriteLine "            <key>Playlist Items</key>"
        fout.WriteLine "            <array>"
        for j = 0 to tracks.Count - 1
          fout.WriteLine "                <dict>"
          fout.WriteLine "                    <key>Track ID</key><integer>" & tracks.Item(j).ID & "</integer>"
          fout.WriteLine "                </dict>"
        next 
        fout.WriteLine "            </array>"
      end if
      fout.WriteLine "        </dict>"
            
      progress.Value = progress.Value + 50
      if Progress.Terminate or Script.Terminate then
        exit for
      end if
    next 
    fout.WriteLine "    </array>"
  end if
  
  fout.WriteLine "</dict>"
  fout.WriteLine "</plist>"
  fout.Close ' Close the output file and finish

  dim ok : ok = not Progress.Terminate and not Script.Terminate
  set Progress = Nothing
  on error resume next
  if not ok then
    fso.DeleteFile(filename) ' remove the output file if terminated
  end if
  SDB.Objects(EXPORTING) = nothing
end sub


sub forcedExport()
  if SDB.Objects(EXPORTING) is nothing then
    Call Export
  end if
end sub


sub ExportITunesXML()
  if SDB.Objects(EXPORTING) is nothing then
    Call Export
  end if
end sub


' Handler for when the Toolbar button is clicked
Sub OnToolbar(btn)
  if SDB.Objects(EXPORTING) is nothing then
    Call Export
  end if
End Sub


' Handler for the timer driving the periodic export
sub periodicExport()
  if getPeriodicExport() and (SDB.Objects(EXPORTING) is nothing) then
    ' if export already in progress silently ignore; otherwise trigger export
    Call Export
  end if
end sub



' Handler for the Export on application shutdown
sub shutdownExport()
  if getExportAtShutdown() and (SDB.Objects(EXPORTING) is nothing) then
    ' if export already in progress silently ignore; otherwise trigger export
    Call Export
  end if
end sub


' Called when MM starts up, installs a timer to export the data
' frequently to the iTunes library.xml.
sub OnStartup

  'MsgBox "DBG: ExportITunesXML OnStartup called"

  Dim btn : Set btn = SDB.Objects("ExportITunesXMLButton")
  If btn Is Nothing Then
    Set btn = SDB.UI.AddMenuItem(SDB.UI.Menu_TbStandard,0,0) 
    btn.Caption = "ExportITunesXML"
    btn.Hint = "Exports all tracks and playlists to an iTunes Music Library.xml file"
    btn.IconIndex = 56
    btn.Visible = True
    Set SDB.Objects("ExportITunesXMLButton") = btn    
  End If
  Call Script.UnRegisterHandler("OnToolbar")
  Call Script.RegisterEvent(btn,"OnClick","OnToolbar")
  ' Option sheet "Library" := -3
  Call SDB.UI.AddOptionSheet("Export to iTunes XML",Script.ScriptPath,"InitSheet","SaveSheet",-3)  
  
  dim exportTimer : set exportTimer = SDB.CreateTimer(3600000) ' export every 60 minutes
  Script.RegisterEvent exportTimer, "OnTimer", "periodicExport"

  Script.RegisterEvent SDB,"OnShutdown","shutdownExport"
end sub


Sub OnInstall()
  'Add entries to script.ini if you need to show up in the Scripts menu
  Dim inip : inip = SDB.ScriptsPath & "Scripts.ini"
  Dim inif : Set inif = SDB.Tools.IniFileByPath(inip)
  If Not (inif Is Nothing) Then
    inif.StringValue("ExportITunesXML","Filename") = "Auto\ExportITunesXML.vbs"
    inif.StringValue("ExportITunesXML","Procname") = "ExportITunesXML"
    inif.StringValue("ExportITunesXML","Order") = "10"
    inif.StringValue("ExportITunesXML","DisplayName") = "Export to iTunes XML"
    inif.StringValue("ExportITunesXML","Description") = "Exports all tracks and playlists to an iTunes library.xml file"
    inif.StringValue("ExportITunesXML","Language") = "VBScript"
    inif.StringValue("ExportITunesXML","ScriptType") = "0"	
	  'inif.StringValue("ExportITunesXML","Shortcut") = "Ctrl+i"
    SDB.RefreshScriptItems
  End If
  Call OnStartup()
End Sub

' Callback to build the configuration dialog
Sub InitSheet(Sheet)
  Dim ini : Set ini = SDB.IniFile  
  Dim ui : Set ui = SDB.UI
  'Dim i : i = 0

	Dim GroupBox0
	Set GroupBox0 = UI.NewGroupBox(Sheet)
	GroupBox0.Caption = "Export to iTunes XML Configuration"
	GroupBox0.Common.SetRect 10, 10, 500, 250

  Dim edt
  Dim y : y = 25


Set edt = ui.NewCheckBox(GroupBox0)
  edt.Common.SetRect 20, y-3, 20, 20
  edt.Common.ControlName = "EITX_ExportAtShutdown"
  edt.Checked = getExportAtShutdown()
  edt.common.Enabled = True
  '
  Set edt = ui.NewLabel(GroupBox0)
  edt.Common.SetRect 40, y, 100, 20
  edt.Caption = "Export at shutdown"
  edt.Autosize = False
  edt.Common.Hint = "If option is set the iTunes library xml will be exported when MediaMonkey is closed. " & _
    "Default is off."
  '
  y = y + 25



  Set edt = ui.NewCheckBox(GroupBox0)
  edt.Common.SetRect 20, y-3, 20, 20
  edt.Common.ControlName = "EITX_PeriodicExport"
  edt.Checked = getPeriodicExport()
  edt.common.Enabled = True
  '
  Set edt = ui.NewLabel(GroupBox0)
  edt.Common.SetRect 40, y, 100, 20
  edt.Caption = "Periodic Export"
  edt.Autosize = False
  edt.Common.Hint = "If option is set the iTunes library xml will be exported every 60 minutes. " & _
    "Default is off."
  '
  y = y + 25


  Set edt = ui.NewLabel(GroupBox0)
  edt.Common.SetRect 20, y+3, 100, 20
  edt.Caption = "Filename:"
  edt.Autosize = False
  edt.Common.Hint = "The file name for the exported iTunes Music Library XML file. " & _
    "If blank/empty the default value of `iTunes Music Library.xml` will be used."
  '
  Set edt = ui.NewEdit(GroupBox0)
  edt.Common.SetRect 80, y, 455-80, 20
  edt.Common.ControlName = "EITX_Filename"
  edt.Text = getFilename()
  edt.common.Enabled = True
  '
  Set edt = ui.NewButton(GroupBox0)
  edt.Common.SetRect 460,y,20,20
  edt.Caption = "..." ' would be nice if we could have a filer icon like in MediaMonkey system dialogs....
  edt.Common.ControlName = "EITX_FileBrowser"    ' to open file browser.... see getExportFilename()
  ' note: selecting a file would also imply setting the directory
  edt.common.Enabled = False ' not yet implemented >> deactivate this control
  
  '
  y = y + 25


  Set edt = ui.NewLabel(GroupBox0)
  edt.Common.SetRect 20, y+3, 100, 20
  edt.Caption = "Directory:"
  edt.Autosize = False
  edt.Common.Hint = "The directory where the iTunes Music Library XML file will be stored. " & _
    "If blank/empty this will be initialised to the default location. On Windows 10 this is typically the `%APPDATA%\MediaMonkey` directory."
  
  Set edt = ui.NewEdit(GroupBox0)
  edt.Common.SetRect 80, y, 455-80, 20
  edt.Common.ControlName = "EITX_Directory"
  edt.Text = getDirectory()
  edt.common.Enabled = True
  '
  Set edt = ui.NewButton(GroupBox0)
  edt.Common.SetRect 460,y,20,20
  edt.Caption = "..." ' would be nice if we could have a folder icon like in MediaMonkey system dialogs....
  edt.Common.ControlName = "EITX_DirectoryBrowser"    ' to open dir browser.... see getExportFilename()
  edt.common.Enabled = False ' not yet implemented >> deactivate this control
  '
  y = y + 25


End Sub

' Callback to store/process when configuration dialog is confimred
Sub SaveSheet(Sheet)
  setExportAtShutdown(Sheet.Common.ChildControl("EITX_ExportAtShutdown").Checked)
  setPeriodicExport(Sheet.Common.ChildControl("EITX_PeriodicExport").Checked)
  '
  setFilename(Sheet.Common.ChildControl("EITX_Filename").Text)
  setDirectory(Sheet.Common.ChildControl("EITX_Directory").Text)
End Sub