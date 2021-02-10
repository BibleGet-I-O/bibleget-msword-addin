# BibleGet plugin for Microsoft Word 2007+

![GitHub release (latest by date)](https://img.shields.io/github/v/release/BibleGet-I-O/bibleget-msword-addin?style=flat-square)
![GitHub code size in bytes](https://img.shields.io/github/languages/code-size/BibleGet-I-O/bibleget-msword-addin?style=flat-square)
![GitHub](https://img.shields.io/github/license/BibleGet-I-O/bibleget-msword-addin?style=flat-square)
![SourceForge](https://img.shields.io/sourceforge/dt/bibleget?style=flat-square)

<table>
  <thead>
    <tr><th colspan=2>About this package</th></tr>
  </thead>
  <tbody>
    <tr><td><b>Author</b></td><td>John Romano D'Orazio</td></tr>
    <tr><td><b>Author email</b></td><td>priest@johnromanodorazio.com</td></tr>
    <tr><td><b>Project Website</b></td><td>https://www.bibleget.io</td></tr>
    <tr><td><b>Latest release</b></td><td>https://sourceforge.net/projects/bibleget/files/latest/download</td></tr>
  </tbody>
</table>

*(The badge download count only takes into account downloads since the package releases were moved to Sourceforge. Add to that the download count from the BibleGet website: 11,967)*

In order to install the plugin, download the most recent installation package from the link here above. Click on the link provided, which will allow you to download the most recent release from the Sourceforge webite. Once downloaded either choose "Execute" from the notification at the bottom of the browser window (MS Explorer or Edge), or click on the downloaded file at the bottom of the browser window (Chrome), or go to your browser downloads (Firefox) and click on the downloaded package from there. If the download gets blocked by your browser, please click on the three dots on the right side of the download notification and choose "Keep anyways" (or similar); if using Edge or Explorer you will have the possibility of reporting the download as safe, please do so if you would like to see this project grow and flourish.

The installation process will launch and will probably take a few minutes, a few Microsoft components need to be downloaded and installed in order for the plugin to work. You may be requested to restart your computer in the process, if that is the case the installation will pick up again and complete after the restart. Once the installation has completed, every time you open Microsoft Word you’ll see a new menu “BibleGet I/O” with it’s own ribbon area. The icon buttons on the ribbon area allow you to set preferences for text formatting of the Bible quotes, see information about the plugin and renew data from the BibleGet server (such as supported Bible Versions, and Languages for the Books of the Bible that the BibleGet server can understand), get instructions for formulating Bible Quotes and using the Add-In, send donations or feedback. There are two ways of inserting Bible quotes into your document: by opening a dialog window where you can choose the Bible versions to quote from and type in your desired quote, or by writing the desired quote directly in your document and selecting it and clicking on the relative icon button. There are keyboard shortcuts for the icon buttons on the ribbon, when you press “ALT” you can see a hint. “ALT+q” is the shortcut for the BibleGet menu, which can then be continued for the single icon items.

This project is released as open source, in the hopes that others might collaborate on the project, or that what I have not succeeded in accomplishing might be picked up again by someone else and made to be of better service to mankind.

This plugin facilitates inserting Bible quotes into your documents.

It communicates with the [BibleGet I/O service endpoint](https://query.bibleget.io) for retrieval of Bible quotes.

You can set your preferred text formatting for the Bible quotes so that you don't have to format them manually every time you insert them into your document.

Credits for the icon images used for the buttons are to be given to https://dryicons.com/icon-packs/wysiwyg-classic .

# Changelog

## Version 3.0.1.8 (October 4, 2020)
* Highlight plain ascii matches other than diacritic matches against search term

## Version 3.0.1.6 (October 2, 2020)
* Allow any kind of emdash, endash, or hyphens as the hyphen separator in query strings
* small fixes on popup windows element alignment
* use png's instead of bmp's for Ribbon icons, with correct background transparency
* make sure that spaces are added between verses when verse numbers are not shown

## Version 3.0.1.3 (September 26, 2020)
* Fix update process using sourceforge releases rather than Wordpress Download Manager

## Version 3.0.1.2 (September 26, 2020)
* Add full word highlighting when partial matches are found when searching verses by keyword: the exact match will be highlighted yellow, the rest of the word will be highlighted light yellow
* Fix verse numbers showing even when verse number visibility is set to hide (thanks to user feedback from Tommaso F.)

## Version 3.0.1.1 (July 8, 2020)
* Fix blocking error on Windows 7 systems which have no internationalization information about the latin language

## Version 3.0.0.7 (July 1, 2020)
* Add search for verses by keyword
* Add layout options to the preferences area
* Fix regular expression issues when text contains special characters
* Use Word interface language for localization rather than system language (they're not necessarily the same)

## Version 2.2.6.0 (March 7, 2016)

* Fixed update process
* Added automatic check for new updates

## Version 2.2.0.0 (February 18, 2016)

* Added debug log functionality to help debug those situations where the AddIn is not working correctly for some reason (debug log created in AppData location on AddIn startup, “enable debug” button added to “Health Status” form), debug log file if existing automatically attached to feedback email
* Added secondary registration of registry keys to make sure keys are installed to actual user registry even in cases where user is not administrator but installation is performed with administrator privileges

## Version 2.1.0.0

* Added internal automatic update check (also added functionality on bibleget.io website that exposes the version information so that the plugin can communicate and obtain this information). Update check scheduled once every 7 days or when Server Data is renewed.

## Version 2.0.0.0

* First release of the BibleGet AddIn for Microsoft Word!
