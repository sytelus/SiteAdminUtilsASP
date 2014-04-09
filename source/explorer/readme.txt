Advanced Remote File Explorer v1.20 (May 30, 2002)

Author: BenWhite@columbus.rr.com

Install:
	Simply extract the contents of the zip file into a directory on your web site

Setup:
	default.asp
		If you plan on viewing this page with ie4 then you need to
		change the scroll property in the body tag to "auto"
		If using ie5 or above, leave the setting = "no"

	dir.asp
		strInitDirectory  - Set the Home Directory for the explorer
		blnDownload       - Enable/Disable Downloads (may be slow for large files)
		blnHidden         - Enable/Disable View of Hidden files/folders
		blnIcon           - Enable/Disable Dynamic Icon Loading (this sometimes fails due to server connection limitations)
		                    The default connection limit for an unlicensed iis server is 10.
		                    There is no way to defeat that setting,
		                    although you can change this setting to 40 or less and it works,
		                    but that is not widely advertised.

	Note:
		If you want to avoid possible download abuse.
		You should delete the download.asp file.
		Remember, blnDownload and blnIcon require this file to work.

Running:
	To use this, simply open a web browser and enter the
	url that points to the path that you extracted the files in.

Adding Icons:
	All icons are 16x16 transparent gif images.
	They follow the naming convention of ext_XXX.gif
	Where XXX = the extension name. (not limited to 3 characters)

	if you think the icons you add would be useful for others,
	please email them to me, and I will include them in the next release.

Known Issues:
	The function used to capture the file link within .lnk files, only work 99% of the time.
	I could not find any microsoft support on how to do this in asp, and was forced
	to created this function by looking at similarities I noticed in the .lnk files.

	Contents of network drives cannot be viewed.
	This is due to user permissions.

	For some versions of internet explorer the download functions will not work properly.
	Microsoft has verified these bugs, and upgrading to the latest version or service pack of ie should fix these issues.
	http://support.microsoft.com/default.aspx?scid=kb;en-us;Q182315
	http://support.microsoft.com/default.aspx?scid=kb;en-us;Q279667
	http://support.microsoft.com/default.aspx?scid=kb;en-us;Q267991
	http://support.microsoft.com/default.aspx?scid=kb;en-us;Q262042
	http://support.microsoft.com/default.aspx?scid=kb;en-us;Q266305
	http://support.microsoft.com/default.aspx?scid=kb;en-us;Q281119
	Thanks to Matthijs Wink for some research on this.

Release Notes:

v1.20 (May 30, 2002)
  Fixed: .lnk files are now displayed 100% correct.
  Added: Alt text on .lnk and .url files now display target of the file on mouseover

v1.19 (May 29, 2002)
  Added: Support for ChilisoftASP (tested on Apache v1.3.24 w/ ChilisoftASP v3.6.2)

v1.18 (May 24, 2002)
  Added: About information
  Fixed: Bug where files with no . would cause an error

v1.17 (May 22, 2002)
  Added: Can now resize the treeview width (25% until user resizes)
  Changed: The displayed path in the title bar reflects initial folder path. (if strInitDirectory <> "")
  Changed: Some code cleanup (more needed)

v1.16 (May 21, 2002)
  Added: More security when home directory is set (strInitDirectory)
  Updated: More efficient way to download files

v1.15 (May 21, 2002)
  Changed: Special Icon folder only appears when entire file system is viewable (strInitDirectory = "")

v1.14 (May 20, 2002)
  Fixed: Object not found bug when click on My Computer icon

v1.13 (May 20, 2002)
  Added: Can now set the initial folder path (limit explorer to sub folders)
  Added: More icons (230 now)

v1.12 (May 15, 2002)
  Removed: ie4 auto-scrolling stuff (messed up ie6)
  Added: More icons (over 200 now)

v1.11 (May 13, 2002)
  Added / Updated: 55 new icons.
  Fixed: Dotted lines now appear correctly when last folder is hidden, and hidden folders not displayed.

v1.10 (May 10, 2002)
  Added: Dynamic ico file loading.
  Added: More support for those poor ie4 users. (scrolling)
  Added: Dotted lines found in Windows Explorer.
  Added: Special registered file type icons folder.
  Added: Support for special windows file types. (desklink, mapimail, cnf, mydocs, pif, shb, shs, xnk)

v1.09 (May 9, 2002)
  Added: Now over 130 file type icons.
  Added: Both file and url shortcuts appear as they would in windows explorer. (lnk file searched)