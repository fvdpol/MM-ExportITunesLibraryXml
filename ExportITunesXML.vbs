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
'       refactor logic to write playlists:
'         - correctly handle playlists with duplicate names
'         - export using same sort order as MediaMonkey
'         - export parent before children (as per iTunes behaviour)
' 1.6.3 reorder xml fields to (better) match iTunes format
'       add Persistent ID for compatibility with Serato DJ
'       add Grouping in export
'       add dummy Library Persistent ID to the header for compatibility with Pioneer Recordbox DJ
'       mark playlists that have sub-playlists as 'folder' (for compatibility with Pioneer Recordbox DJ)
'       add "Play Date" (timestamp in numeric format) in addition to the "Play Date UTC"
' 1.6.4 add feature/option to exclude the playlist section in the generated xml file
'       add DebugMsg() function and support framework
'       suppress Anti Malware Scan Interface AMSI_ATTRIBUTE_CONTENT_NAME Error 0x80070490 being raised
'       resizable Options dialog
'       add file and directory browser in the Options dialog
'       restructure Options dialog, create logical grouping for settings
'
'
option explicit     ' report undefined variables, ...

Dim Debug
Debug = getDebug()
setDebug(Debug) ' write the flag in config file (for manual easy changing)
'  ------------------------------------------------------------------
const EXPORTING = "itunes_export_active"
dim scriptControl

sub DebugMsg(ByVal myMsg)
  if Debug then SDB.Tools.OutputDebugString("ExportITunesXML: " & myMsg)
end sub

' Returns encoded URI for provided location string.
function encodeLocation(ByVal location)
  ' 10.10.2010: need jscript engine to access its encodeURI function which is not
  ' available in vbscript

  if isEmpty(scriptControl) then
    set scriptControl = CreateObject("ScriptControl")
    scriptControl.Language = "JScript"
    '
    ' running the JScript function (scriptControl.Run()) results in an error being logged in DbgView
    ' [14856] [2018-11-18 18:17:37.828] [error  ] [AMSI       ] [14856: 6220] AMSI_ATTRIBUTE_CONTENT_NAME Error 0x80070490
    ' >> AMSI is the Windows "Anti Malware Scan Interface"; used by Windows Defender and AVG
    ' setting the AllowUI to false seems to prevent these errors from being raised (why is that???)
    scriptControl.AllowUI = False
  end if
  location = replace(location, "\", "/")
  encodeLocation = scriptControl.Run("encodeURI", location)
  encodeLocation = replace(encodeLocation, "#", "%23")    ' # is not permitted in path
  encodeLocation = replace(encodeLocation, "&", "&#38;")  ' amparsant needs to be escaped
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
function escapeXML(ByVal srcstring)
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


' Getter for the configured Debug boolean
function getDebug()
  dim myIni
  dim myValue
  dim myBool

  set myIni = SDB.IniFile
  myValue = cleanFilename(myIni.StringValue("ExportITunesXML","Debug"))

  ' parse ini value to boolean; use default if not defined as 0/1
  if myValue = "0" then
    myBool = False
  elseif myValue = "1" then
    myBool = True
  else
    myBool = getDefaultDebug()
  end if

  getDebug = myBool
end function
'
' Setter for the configured Debug boolean
sub setDebug(byVal myBool)
  dim myIni
  set myIni = SDB.IniFile

  if myBool then
    myIni.StringValue("ExportITunesXML","Debug") = "1"
  else
    myIni.StringValue("ExportITunesXML","Debug") = "0"
  end if
end sub
'
function getDefaultDebug()
  getDefaultDebug = False
end function


' Getter for the configured ExportAtShutdown boolean
function getExportAtShutdown()
  dim myIni
  dim myValue
  dim myBool

  set myIni = SDB.IniFile
  myValue = cleanFilename(myIni.StringValue("ExportITunesXML","ExportAtShutdown"))

  ' parse ini value to boolean; use default if not defined as 0/1
  if myValue = "0" then
    myBool = False
  elseif myValue = "1" then
    myBool = True
  else
    myBool = getDefaultExportAtShutdown()
  end if

  getExportAtShutdown = myBool
end function
'
' Setter for the configured ExportAtShutdown boolean
sub setExportAtShutdown(byVal myBool)
  dim myIni
  set myIni = SDB.IniFile

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

  ' parse ini value to boolean; use default if not defined as 0/1
  if myValue = "0" then
    myBool = False
  elseif myValue = "1" then
    myBool = True
  else
    myBool = getDefaultPeriodicExport()
  end if

  getPeriodicExport = myBool
end function
'
' Setter for the configured PeriodicExport boolean
sub setPeriodicExport(byVal myBool)
  dim myIni
  set myIni = SDB.IniFile

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
  dim myIni
  dim myDirectory

  set myIni = SDB.IniFile
  myDirectory = cleanDirectoryName(myIni.StringValue("ExportITunesXML","Directory"))

  if isValidDirectory(myDirectory) = False then
    myDirectory = getDefaultDirectory()
  end if

  getDirectory = myDirectory
end function

' Setter for the configured Directory
sub setDirectory(byVal myDirectory)
  dim myIni
  myDirectory = cleanDirectoryName(myDirectory)

  if isValidDirectory(myDirectory) = False then
    myDirectory = getDefaultDirectory()
  end if

  set myIni = SDB.IniFile
  myIni.StringValue("ExportITunesXML","Directory") = myDirectory

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

function cleanDirectoryName(byVal myDirectory)
  ' do so basic cleanup; ensure the path ends with a directory separator
  if right(myDirectory,1) <> "\" then
    ' simply append the missing separator
    myDirectory = myDirectory & "\"
  end if
  cleanDirectoryName = myDirectory
end function

' Return true if the directory is defined
function isValidDirectory(byVal myDirectory)
  dim myResult : myResult = True

  ' check for blank/empty directory
  if trim(myDirectory) = "" or trim(myDirectory) = "\" then
    myResult = False
  end if

  ' potential test to check if the directory is actually writable...
  '
  ' for now assume all will be good and trap any errors writing to the
  ' in the actual export routine.

  isValidDirectory = myResult
end function


' Getter for the configured Filename
' if filename is undefined/blank then return the default
function getFilename()
  dim myIni
  dim myFilename

  set myIni = SDB.IniFile
  myFilename = cleanFilename(myIni.StringValue("ExportITunesXML","Filename"))

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

   if myFilename = "" then
    myFilename = getDefaultFilename()
  end if

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


' Getter for the configured NoPlaylistExport boolean
function getNoPlaylistExport()
  dim myIni
  dim myValue
  dim myBool

  set myIni = SDB.IniFile
  myValue = cleanFilename(myIni.StringValue("ExportITunesXML","NoPlaylistExport"))

  ' parse ini value to boolean; use default if not defined as 0/1
  if myValue = "0" then
    myBool = False
  elseif myValue = "1" then
    myBool = True
  else
    myBool = getDefaultNoPlaylistExport()
  end if

  getNoPlaylistExport = myBool
end function
'
' Setter for the configured ExportAtShutdown boolean
sub setNoPlaylistExport(byVal myBool)
  dim myIni
  set myIni = SDB.IniFile

  if myBool then
    myIni.StringValue("ExportITunesXML","NoPlaylistExport") = "1"
  else
    myIni.StringValue("ExportITunesXML","NoPlaylistExport") = "0"
  end if
end sub
'
function getDefaultNoPlaylistExport()
  getDefaultNoPlaylistExport = False
end function



' N must be numberic. Return value is N converted to a string, padded with
' a single "0" if N has only one digit.
function LdgZ(ByVal N)
  if (N >= 0) and (N < 10) then
    LdgZ = "0" & N
  else
    LdgZ = "" & N
  end if
end function

' Adds a simple key/value pair to the XML accessible via textfile fout.
sub addKey(ByVal fout, ByVal key, ByVal val, ByVal keytype)
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


function getExportFilename()
  getExportFilename = getDirectory() + getFilename()
end function


Function ConvertToUnixTimeStamp(byVal myDateTime)
 Dim d : d = CDate(myDateTime)
 ConvertToUnixTimeStamp = DateDiff("s", "01/01/1970 00:00:00", d)
End Function


' similar to the unix timestamp, but then seconds since 1 jan 1904 (whattah????)
Function ConvertToItunesIntegerTimeStamp(byVal myDateTime)
 Dim d : d = CDate(myDateTime)
 ConvertToItunesIntegerTimeStamp = DateDiff("s", "01/01/1904 00:00:00", d)
End Function


' find the parent playlist ID for given playlist. Returns 0 if no parent exists
function getparentID(byVal myPlaylist)
	dim iter
  dim myPlaylistID, myParentID
	myPlaylistID = myPlaylist.ID
  myParentID = 0
  set iter = SDB.Database.OpenSQL("select ParentPlaylist from PLAYLISTS where IDPlaylist=" & myPlaylistID)
  while not iter.EOF
    myParentID = iter.ValueByIndex(0)
    iter.next
  wend
  set iter = nothing
	getparentID = myParentID
end function



' process one level of playlist; traverse into child playlists where needed
sub WritePlaylist(fout, progress, byval progressText, byval myPlaylist)
  dim myChildPlaylists : Set myChildPlaylists = myPlaylist.ChildPlaylists
  dim i,j, playlist
  dim parentID
  dim tracks

  for i = 0 To myChildPlaylists.Count - 1                             ' For all (first-level) playlists in List...
    Set playlist = myChildPlaylists.Item(i)                                ' ... print out the number of child playlists and tracks
	  parentID = getparentID(playlist)
    set tracks = playlist.Tracks

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
    fout.WriteLine "            <key>All Items</key><true/>"

    ' if this playlist has any childs playlists add the Folder=true flag to have Pioneer Recordbox correctly parse the folder structure
    if (playlist.ChildPlaylists.count > 0) then
			fout.WriteLine "            <key>Folder</key><true/>"
    end if

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

    ' if this playlist has any childs playlists traverse through them...
    if (playlist.ChildPlaylists.count > 0) then
      call WritePlaylist(fout, progress, progressText, playlist)
    end if
  next
end sub



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

  DebugMsg("Export starting...")

  dim filename, fso, iter, songCount, fout, progress, song, playlistCount
  dim progressText, i, j, tracks, playlist

  filename = getExportFilename()
  if filename = "" then
    SDB.Objects(EXPORTING) = nothing
    exit sub
  end if

  set fso = SDB.Tools.FileSystem
  set fout = fso.CreateTextFile(filename, true)

  if fout is nothing then
     MsgBox SDB.Localize("Unable to write to '" & filename & "'."), 64, "iTunes Export Script"

     ' cleanup
     set fso = nothing
     SDB.Objects(EXPORTING) = nothing
    exit sub
  end if

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
  fout.WriteLine "    <key>Library Persistent ID</key><string>null</string>" ' add dummy to keep Recordbox DJ Happy
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
    DebugMsg("Track Export...")
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
      addKey fout, "Size", Song.FileLength, "integer"
      addKey fout, "Total Time", Song.SongLength, "integer"
      if Song.DiscNumber >= 0 then addKey fout, "Disc Number", Song.DiscNumber, "integer" ' potential type problem with DiscNumberStr
      ' Field not available in MM: <key>Disc Count</key>
      if Song.TrackOrder >= 0 then addKey fout, "Track Number", Song.TrackOrder, "integer" ' potential type problem with TrackOrderStr
      ' Field not available in MM: <key>Track Count</key>
      if Song.Year > 0 then addKey fout, "Year", Song.Year, "integer"
      if Song.BPM > 0 then addKey fout, "BPM", Song.BPM, "string"
      addKey fout, "Date Modified", Song.FileModified, "date"
      addKey fout, "Date Added", Song.DateAdded, "date"
      addKey fout, "Bit Rate", Int(Song.Bitrate / 1000), "integer"
      addKey fout, "Sample Rate", Song.SampleRate, "integer"
      if Song.PlayCounter > 0 then addKey fout, "Play Count", Song.PlayCounter, "integer"
      if Song.LastPlayed > 0 then
        addKey fout, "Play Date", ConvertToItunesIntegerTimeStamp(Song.LastPlayed), "integer"
        addKey fout, "Play Date UTC", Song.LastPlayed, "date"
      end if
      ' Field not available: <key>Skip Count</key><integer>1</integer>
			' Field not available: <key>Skip Date</key><date>2018-10-19T16:06:26Z</date>
			if Song.Rating >= 0 and Song.Rating <= 100 then
        addKey fout, "Rating", Song.Rating, "integer" ' rating seems to be compatible in range (although not stored in same id3 tag)
      end if

			' Field not available: <key>Loved</key><true/>
			' Field not available: <key>Compilation</key><true/>

      addKey fout, "Persistent ID", Song.id, "string"   ' Field not available in MM, but simulate this as Serato needs this field
      addKey fout, "Track Type", escapeXML("File"), "string"
      addKey fout, "File Folder Count", -1, "integer"
      addKey fout, "Library Folder Count", -1, "integer"
      addKey fout, "Name", escapeXML(Song.Title), "string"
      if Song.ArtistName <> "" then addKey fout, "Artist", escapeXML(Song.ArtistName), "string"
      if Song.AlbumArtistName <> "" then addKey fout, "Album Artist", escapeXML(Song.AlbumArtistName), "string"
      if Song.MusicComposer <> "" then addKey fout, "Composer", escapeXML(Song.MusicComposer), "string"
      if Song.AlbumName <> "" then addKey fout, "Album", escapeXML(Song.AlbumName), "string"
      if Song.Grouping <> "" then addKey fout, "Grouping", escapeXML(Song.Grouping), "string"
      if Song.Genre <> "" then addKey fout, "Genre", escapeXML(Song.Genre), "string"
      addKey fout, "Kind", escapeXML("MPEG audio file"), "string"
      if Song.Comment <> "" then addKey fout, "Comments", escapeXML(Song.Comment), "string"

      addKey fout, "Location", encodeLocation("file://localhost/" & Song.Path), "string"

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
  if playlistCount > 0 and getNoPlaylistExport() = False and not Progress.Terminate and not Script.Terminate then
    DebugMsg("Playlist Export...")
    fout.WriteLine "    <key>Playlists</key>"
    fout.WriteLine "    <array>"

    Dim RootPlaylist : Set RootPlaylist = SDB.PlaylistByID(-1) ' Playlist represents the root (virtual) playlist
    call WritePlaylist(fout, progress, progressText, RootPlaylist)

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

  DebugMsg("Export finished")
end sub


sub ExportITunesXML()
  if SDB.Objects(EXPORTING) is nothing then
    Call Export
  end if
end sub


' Handler for when the Toolbar button is clicked
Sub OnToolbar(myButton)
  if SDB.Objects(EXPORTING) is nothing then
    Call Export
  end if
End Sub


' Handler for the timer driving the periodic export
sub periodicExport(myTimer)
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


' Called when MM starts up
sub OnStartup
  DebugMsg("OnStartup()")
  ' Create and register toolbar button
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

  ' Register Option sheet as child under "Library" := -3
  'myOptionSheet = SDB.UI.AddOptionSheet("Export to iTunes XML",Script.ScriptPath,"InitSheet","SaveSheet",-3)
  Dim myOptionSheet
  myOptionSheet = SDB.UI.AddOptionSheetEx("Export to iTunes XML",Script.ScriptPath,"InitSheet","SaveSheet","CancelSheet",-3)
  
  ' Register handler for the periodic export
  dim exportTimer : set exportTimer = SDB.CreateTimer(3600000) ' export every 60 minutes (arg in ms)
  Set SDB.Objects("ExportITunesXMLExportTimer") = exportTimer
  Script.RegisterEvent exportTimer, "OnTimer", "periodicExport"

  ' Register handler for the export on shutdown
  Script.RegisterEvent SDB,"OnShutdown","shutdownExport"
  DebugMsg("OnStartup() finished")
end sub


Sub OnInstall()
  DebugMsg("OnInstall()")
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
  DebugMsg("OnInstall() finished")
End Sub

' Callback to build the configuration dialog
Sub InitSheet(Sheet)
  Dim y : y = 0

  SDB.Objects("ConfigSheet") = Sheet  
  Sheet.Common.ControlName = "EITX_ConfigSheet"
  Script.RegisterEvent Sheet.Common, "OnResize", "ResizeSettingSheet"

  y = y + CreateGroupbox0(Sheet, y, "Output File")
  y = y + CreateGroupbox1(Sheet, y, "Automatic Export")
  y = y + CreateGroupbox2(Sheet, y, "Output Modification")

  ' force resize event to make the layout consistent
  Call ResizeSettingSheet(Sheet)

End Sub


' create groupbox for the output selector
' returns height
function CreateGroupbox0(Sheet, Top, Caption)
  Dim ui : Set ui = SDB.UI
  Dim box, lbl, edt, btn
  Dim y : y = 0

	Set box = UI.NewGroupBox(Sheet)
  box.Common.ControlName = "EITX_Groupbox0"
	box.Caption = Caption
	box.Common.SetRect 10, Top + 10, 500, 200
  y = y + 25


  Set lbl = ui.NewLabel(box)
  lbl.Common.SetRect 20, y+3, 100, 20
  lbl.Caption = "Filename:"
  lbl.Autosize = False
  lbl.Common.Hint = "The file name for the exported iTunes Music Library XML file. " & _
    "If blank/empty the default value of `iTunes Music Library.xml` will be used."
  '
  Set edt = ui.NewEdit(box)
  edt.Common.SetRect 80, y, 455-80, 20
  edt.Common.ControlName = "EITX_Filename"
  edt.Text = getFilename()
  edt.Common.Hint = lbl.Common.Hint
  edt.common.Enabled = True
  '
  Set btn = ui.NewButton(box)
  btn.Common.SetRect 460,y,20,20
  btn.Caption = "..." ' would be nice if we could have a filer icon like in MediaMonkey system dialogs....
  btn.Common.ControlName = "EITX_FileBrowser"
  btn.Common.Hint = "Click to open File Browser"
  btn.common.Enabled = True
  Call Script.RegisterEvent(btn.common, "OnClick", "FileBrowser") ' note: selecting a file would also imply setting the directory
  '
  y = y + 25


  Set lbl = ui.NewLabel(box)
  lbl.Common.SetRect 20, y+3, 100, 20
  lbl.Caption = "Directory:"
  lbl.Autosize = False
  lbl.Common.Hint = "The directory where the iTunes Music Library XML file will be stored. " & _
    "If blank/empty this will be initialised to the default location. On Windows 10 this is typically the `%APPDATA%\MediaMonkey\` directory."
  '
  Set edt = ui.NewEdit(box)
  edt.Common.SetRect 80, y, 455-80, 20
  edt.Common.ControlName = "EITX_Directory"
  edt.Text = getDirectory()
  edt.Common.Hint = lbl.Common.Hint
  edt.common.Enabled = True
  '
  Set btn = ui.NewButton(box)
  btn.Common.SetRect 460,y,20,20
  btn.Caption = "..." ' would be nice if we could have a folder icon like in MediaMonkey system dialogs....
  btn.Common.ControlName = "EITX_DirectoryBrowser"
  btn.Common.Hint = "Click to open Directory Browser"
  btn.common.Enabled = True 
  Call Script.RegisterEvent(btn.common, "OnClick", "DirectoryBrowser")
  '
  y = y + 25


  ' end of groupbox
  box.Common.Height = y
  y = y + 5

  CreateGroupbox0 = y
end function


' create groupbox for the schedule
' returns height
function CreateGroupbox1(Sheet, Top, Caption)
  Dim ui : Set ui = SDB.UI
  Dim box, lbl, edt, btn
  Dim y : y = 0

  Set box = UI.NewGroupBox(Sheet)
  box.Common.ControlName = "EITX_Groupbox1"
	box.Caption = Caption
	box.Common.SetRect 10, Top + 10, 500, 200
  y = y + 25


  Set edt = ui.NewCheckBox(box)
  edt.Common.SetRect 20, y-3, 20, 20
  edt.Common.ControlName = "EITX_ExportAtShutdown"
  edt.Checked = getExportAtShutdown()
  edt.common.Enabled = True
  '
  Set edt = ui.NewLabel(box)
  edt.Common.SetRect 40, y, 100, 20
  edt.Caption = "Export at shutdown"
  edt.Autosize = False
  edt.Common.Hint = "If option is set the iTunes library xml will be exported when MediaMonkey is closed. " & _
    "Default is off."
  y = y + 25


  Set edt = ui.NewCheckBox(box)
  edt.Common.SetRect 20, y-3, 20, 20
  edt.Common.ControlName = "EITX_PeriodicExport"
  edt.Checked = getPeriodicExport()
  edt.common.Enabled = True
  '
  Set edt = ui.NewLabel(box)
  edt.Common.SetRect 40, y, 100, 20
  edt.Caption = "Periodic Export"
  edt.Autosize = False
  edt.Common.Hint = "If option is set the iTunes library xml will be exported every 60 minutes. " & _
    "Default is off."
  y = y + 25


' end of groupbox
  box.Common.Height = y
  y = y + 5

  CreateGroupbox1 = y
end function


' create groupbox for the output modifications
' returns height
function CreateGroupbox2(Sheet, Top, Caption)
  Dim ui : Set ui = SDB.UI
  Dim box, lbl, edt, btn
  Dim y : y = 0

  Set box = UI.NewGroupBox(Sheet)
  box.Common.ControlName = "EITX_Groupbox2"
	box.Caption = Caption
	box.Common.SetRect 10, Top + 10, 500, 200
  y = y + 25
  
  
  Set edt = ui.NewCheckBox(box)
  edt.Common.SetRect 20, y-3, 20, 20
  edt.Common.ControlName = "EITX_NoPlaylistExport"
  edt.Checked = getNoPlaylistExport()
  edt.common.Enabled = True
  '
  Set edt = ui.NewLabel(box)
  edt.Common.SetRect 40, y, 100, 20
  edt.Caption = "Exclude export of Playlists"
  edt.Autosize = False
  edt.Common.Hint = "If option is set the iTunes library xml will only contain the tracks; the playlists will be excluded. " & _
    "Default is off."
  '
  y = y + 25
  

' end of groupbox
  box.Common.Height = y
  y = y + 5

  CreateGroupbox2 = y
end function





' note: selecting a file would also imply setting the directory
Sub FileBrowser(Obj)
  Dim Sheet 
  Set Sheet = SDB.Objects("ConfigSheet")
  If Not (Sheet is Nothing) then
    Dim edtFile, edtDir
    Set edtFile = Sheet.Common.ChildControl("EITX_Filename")
    Set edtDir = Sheet.Common.ChildControl("EITX_Directory")
    If Not (edtDir is Nothing or edtDir is Nothing) then      
      ' Create common dialog and ask where to save the file   
      ' https://www.mediamonkey.com/wiki/index.php?title=SDBCommonDialog
      Dim dlg
      Set dlg = SDB.CommonDialog
      dlg.DefaultExt="xml"
      dlg.Filter = "XML Document (*.xml)|*.xml|All files (*.*)|*.*"
      dlg.FilterIndex = 1
      dlg.Flags=cdlOFNOverwritePrompt + cdlOFNHideReadOnly
      dlg.InitDir = edtDir.Text
      dlg.ShowSave

      If Not dlg.Ok Then
        Exit Sub   ' if cancel was pressed, exit
      End If

      ' Get the selected filename
      Dim FullName
      FullName = dlg.FileName
      'DebugMsg("got fullname = " & FullName)
      Dim pos, NewDir, NewFilename
      pos = InStrRev(FullName, "\")
      If pos > 0 Then
        NewDir = Left(FullName, pos)
        NewFileName = Mid(FullName, pos+1)
      Else
        NewDir = ""
        NewFilename = FullName
      End If
      'DebugMsg("got dir  = " & NewDir)
      'DebugMsg("got file = " & NewFilename)
      edtDir.Text = NewDir
      edtFile.Text = NewFilename

     End if
  End if 
End Sub


Sub DirectoryBrowser(Obj)
  Dim Sheet 
  Set Sheet = SDB.Objects("ConfigSheet")
  If Not (Sheet is Nothing) then
    Dim edt 
    Set edt = Sheet.Common.ChildControl("EITX_Directory")
    If Not (edt is Nothing) then      
      Dim str 
      str = SDB.SelectFolder(edt.Text,"Select Directory Path for storing the iTunes Library")
      If Not (str = "") Then
        If Right(str,1) = "\" Then
          edt.Text = str
        Else
          edt.Text = str & "\"
        End If
      End If
    End if
  End if 
End Sub


Sub ResizeSettingSheet(Control)
  Dim FrameWidth 
  Dim ctrl

'  DebugMsg("ResizeSettingSheet()")
'  DebugMsg("ControlName = " & Control.Common.ControlName)
'  DebugMsg("Width = " & Control.Common.Width)

  FrameWidth = Control.Common.Width
  Dim Sheet : Set Sheet = Control.Common.TopParent 
  
  Set ctrl = Sheet.Common.ChildControl("EITX_Groupbox0")
  ctrl.Common.Width = FrameWidth - 20

  Set ctrl = Sheet.Common.ChildControl("EITX_Filename")
  ctrl.Common.Width = FrameWidth - 20 - 35 - 80

  Set ctrl = Sheet.Common.ChildControl("EITX_FileBrowser")
  ctrl.Common.Left = FrameWidth - 20 - 30

  Set ctrl = Sheet.Common.ChildControl("EITX_Directory")
  ctrl.Common.Width = FrameWidth - 20 - 35 - 80

  Set ctrl = Sheet.Common.ChildControl("EITX_DirectoryBrowser")
  ctrl.Common.Left = FrameWidth - 20 - 30

  Set ctrl = Sheet.Common.ChildControl("EITX_Groupbox1")
  ctrl.Common.Width = FrameWidth - 20

  Set ctrl = Sheet.Common.ChildControl("EITX_Groupbox2")
  ctrl.Common.Width = FrameWidth - 20

End Sub

' remove any eventhandlers/objects associated with the Configuration Sheet
Sub DisposeSheet()
  Dim Sheet : Set Sheet = SDB.Objects("ConfigSheet")
  'DebugMsg("DisposeSheet()")
  if Not (Sheet is Nothing) Then
    Script.UnRegisterEvents Sheet.Common

    dim btn
    Set btn = Sheet.Common.ChildControl("EITX_FileBrowser")
    if Not (btn is Nothing) then
      Script.UnRegisterEvents btn.Common
    end if
    '
    Set btn = Sheet.Common.ChildControl("EITX_DirectoryBrowser")
    if Not (btn is Nothing) then
      Script.UnRegisterEvents btn.Common
    end if

    SDB.Objects("ConfigSheet") = Nothing
  end If
end Sub

' Callback to store/process when configuration dialog is confimred
Sub SaveSheet(Sheet)
  'DebugMsg("SaveSheet")
  '
  setExportAtShutdown(Sheet.Common.ChildControl("EITX_ExportAtShutdown").Checked)
  setPeriodicExport(Sheet.Common.ChildControl("EITX_PeriodicExport").Checked)
  '
  setFilename(Sheet.Common.ChildControl("EITX_Filename").Text)
  setDirectory(Sheet.Common.ChildControl("EITX_Directory").Text)
  '
  setNoPlaylistExport(Sheet.Common.ChildControl("EITX_NoPlaylistExport").Checked)

  Call DisposeSheet()
End Sub

Sub CancelSheet(Sheet)
  Call DisposeSheet()
End Sub
