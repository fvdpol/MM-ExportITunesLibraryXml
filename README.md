# MediaMonkey Export to iTunes library.xml


Based on the original script posted by "DC" on the MediaMonkey forum to export to SqueezeCenter:

http://www.mediamonkey.com/forum/viewtopic.php?f=2&t=31680


My primary use-case for this script is to export the MediaMonkey library/playlists for use in Native Instruments TraktorDJ


## Installation

The script needs to be registered in MediaMonkey to make it appear in the Reports menu.
In MM\Scripts\scripts.ini, append the following lines:

``` 
[ExportITunes]
FileName=Auto\Export to iTunes library-xml.vbs
ProcName=Export
Order=5
DisplayName=Tracks and Playlists (&iTunes library.xml)
Description=Exports all tracks and playlists to an iTunes library.xml file
Language=VBScript
ScriptType=1
```


## History

Script Original Version; by DC posted Wed Aug 06, 2008 3:01 pm
http://www.mediamonkey.com/forum/viewtopic.php?f=2&t=31680#p162175

Update for Unicode ; by DC posted Sun Oct 10, 2010 6:21 pm
http://www.mediamonkey.com/forum/viewtopic.php?f=2&t=31680&start=60#p272344


Update to add BPM field for Traktor; by Rhashime posted Sat Dec 24, 2011 12:14 pm
http://www.mediamonkey.com/forum/viewtopic.php?f=2&t=31680&start=60#p324753


Update to export playlist structure to Traktor; by Mazze_HH posted Wed Dec 12, 2012 3:51 am
http://www.mediamonkey.com/forum/viewtopic.php?f=2&t=31680&start=60#p354155
