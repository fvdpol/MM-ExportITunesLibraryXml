# MediaMonkey Export to iTunes library.xml



Based on the original script posted by "DC" on the MediaMonkey forum to export to SqueezeCenter:

http://www.mediamonkey.com/forum/viewtopic.php?f=2&t=31680


Primary use-case for this script is to export the MediaMonkey library/playlists for use in Native Instruments TraktorDJ: as bridge to DJ software.


### Tested Versions:

| Software                      | Version   | OS          | Note                                                          |
|-------------------------------|-----------|-------------|---------------------------------------------------------------|
| Media Monkey Gold             | 4.1.20.1864 4.1.21.1875 | Windows 10 | OK                                                                     |
| Native Instruments Traktor DJ | 2.11.3.17 | Windows 10  | OK                                                                      |
| Native Instruments Traktor DJ | 2.11.3.17 | MacOS Sierra | tracks / playlists read OK, file location needs to be updated/remapped |
| Mixxx                         | 2.0.0_x64 2.1.3_x64 2.1.4_x64 | Windows 10  | OK |
| Virtual DJ                    | 8.3 b4459 (2018) | Windows 10 | OK |
| Serato DJ Pro                 | 2.0.3 + | Windows 10 | OK   |
| Pioneer Recordbox DJ          | 5.4.0  |  Windows 10  | OK  |
| Plex                          |        |              | file location needs to be updated/remapped if Plex is running on Linux |
| Musicbee                      |        |              | OK |



Reports on compatibility/issues with other software welcome.


## Installation

Install the ExportITunesXML.mmip as MediaMonkey extension.

Build the binary from sources or download from the repository:
https://github.com/fvdpol/MM-ExportITunesLibraryXml/releases/latest


Default file location for the generated "iTunes Music Library.xml" file is in same location where the MediaMonkey database is stored. On Windows 10 this is typically the `%APPDATA%\MediaMonkey` directory.

## Configuration

A number of settings that were initially harcoded / configured in the script can (since version 1.6.2) be managed via the MediaMonkey Options dialog. 

Navigate to Tools menu -> Options, and open the "Export to iTunes XML configuration dialog within the Library section. 

| Setting            | Description |
|--------------------|-------------|
| Export at Shutdown | If option is set the iTunes library xml will be exported when MediaMonkey is closed. <br> Default is off.| 
| Periodic Export    | If option is set the iTunes library xml will be exported every 60 minutes. <br> Default is off.|
| Filename           | The file name for the exported iTunes Music Library XML file. <br> If blank/empty the default value of `iTunes Music Library.xml` will be used.|
| Directory          | The directory where the iTunes Music Library XML file will be stored. <br> If blank/empty this will be initialised to the default location. On Windows 10 this is typically the `%APPDATA%\MediaMonkey` directory. |
| Exclude export of Playlists | If option is set the iTunes library xml will only contain the tracks; the playlists will be excluded. <br> Default is off.|


Note: Serato expects the xml file to be available in the original location where iTunes stores the file, which is typically in `C: \Users\{user}\Music\iTunes`

## Nested playlists, Folders and Traktor

Apple iTunes and MediaMonkey handle nested playlists / folders in a different way. Due to this difference the result in applications like Native Instruments Traktor may sometimes not be what one expects.

### iTunes:
- Playlists can be 'manual' playlists or auto playlist using some rules -- similar to MediaMonkey
- Folders are a special object that can contain playlists or folders. in the XML the folder contains the contents of all children

### MediaMonkey:

- Playlists can be 'manual' playlists or auto playlist using some rules -- similar to iTunes
- Playlists can be nested: a playlist can contain tracks (as per above) but also child playlists

The main difference is that in iTunes a folder will always show the contents of all folders and/or playlists underneath, while MediaMonkey can have parent playlists that are empty or contain something different. 


### Workaround

To emulate the iTunes folders in MediaMonkey following construct can be applied in MediaMonkey:
- Create an Auto Playlist that will serve as folder
- Move/add the child playlists under this Auto Playlist
- Set the rule (advanced tab) 'is playlist' and checkmark the child playlists

If you select the "Folder" Auto Playlist you should see all tracks from the child playlists. Traktor should recognise this as folder and show folder with the child playlists underneath.


Note that Native Instruments Traktor is known to ignore / filter-out playlists that do not contain any tracks. A work-around could be to include a dummy track.


## Version History

### Version 1.6.4
_in progress_
- add feature/option to exclude the playlist section in the generated xml file


### Version 1.6.3
_Released on November 16, 2018_
- reorder xml fields to (better) match iTunes format
- add Persistent ID for compatibility with Serato DJ
- add Grouping in export
- add dummy Library Persistent ID to the header for compatibility with Pioneer Recordbox DJ
- mark playlists that have sub-playlists as 'folder' (for compatibility with Pioneer Recordbox DJ)
- add Play Date in Apple Numeric format (seconds since 1/1/1904)


### Version 1.6.2
_Released on September 17, 2018_
- Added Options dialog
- Dynamically configurable options for export at shutdown and periodic export
- Dynamically configurable filename and directory
- Refactor logic to write playlists:
    - correctly handle playlists with duplicate names 
    - export using same sort order as MediaMonkey
    - export parent before children (as per iTunes behaviour)


### Version 1.6.1
_Released on August 25, 2018_
- Improved utf-8 unicode handling; support for utf-16 surrogate pairs


### Version 1.6
_Released on june 20, 2018_
- Migrated report to a standard script
- Added MediaMonkey package installer (ExportItunesXML.mmip)
- Added auto-update of script from GitHub repository


### Version 1.5
_Released on June 14, 2018_
- Update to add BPM field for Traktor; by Rhashime posted Sat Dec 24, 2011 12:14 pm
http://www.mediamonkey.com/forum/viewtopic.php?f=2&t=31680&start=60#p324753

- added forced export on shutdown; by Mazze_HH posted Wed Dec 12, 2012 3:51 am
- Update to export playlist structure to Traktor; by Mazze_HH posted Wed Dec 12, 2012 3:51 am
http://www.mediamonkey.com/forum/viewtopic.php?f=2&t=31680&start=60#p354155


### Version 1.4
_Released on June 6, 2018_
- Fixed: Traktor failing import due to invalid characters in xml (& -> `&#38;`); fvdpol, July 4 2018


### Version 1.3
_Released on June 6, 2018_
- Fixed URI encoding, added Last Player; by VariableFlame posted Sun Jan 07, 2018 5:10 pm
http://www.mediamonkey.com/forum/viewtopic.php?f=2&t=31680&start=75#p441991


### Version 1.2
_Released on June 6, 2018_
- Update for Unicode ; by DC posted Sun Oct 10, 2010 6:21 pm
http://www.mediamonkey.com/forum/viewtopic.php?f=2&t=31680&start=60#p272344


### Version 1.0
_Released on June 6, 2018_
- Script Original Version; by DC posted Wed Aug 06, 2008 3:01 pm
http://www.mediamonkey.com/forum/viewtopic.php?f=2&t=31680#p162175


## Future changes/enhancements/ideas

- configuration settings for file location & flags (auto export, periodic export) 
- if needed: configurable quirks for compatibility with other appliciations
- document compatibility with other applications/dj software (feedback from users required)
- selection of playlists to be exported 

