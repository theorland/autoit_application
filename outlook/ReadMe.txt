How to "install" the OutlookEX UDF				2014-02-08
--------------------------------------------------------------------------
This step is mandatory
* Copy OutlookEX.au3 and OutlookEXConstants.au3 into one of the following directories:
  * %ProgramFiles%\AutoIt3\Include
  * Directory as defined in SciTe -> Tools- SciTEConfig -> General Settings -> AutoIt3 Directory Settings -> User Includes
  * The directory where your scripts are located

For SciTE integration (user calltips and syntax highlighting) run SciTe -> Tools- SciTEConfig -> Other tools -> Run User CallTip Manager
to create files au3.user.calltips.api and au3.userudfs.properties in directory %Userprofile% or copy both files from the ZIP-file to the
directory. If they already exist then add the content from the ZIP-files to the already existing files.
For details please check: http://www.autoitscript.com/wiki/Adding_UDFs_to_AutoIt_and_SciTE

Help files and examples
* Copy the *.htm and the remaining *.au3 files to any directory you like. 
  You can't call the help and example scripts from the AutoIt help at the moment


How to use the OutlookEX UDF 		                        2011-05-18
--------------------------------------------------------------------------
* Every script has to have the following format:
  _OL_Open()	                 ; open a onnection to Outlook
  calls to other _OL-functions   ; query or manipulate Outlook items
  _OL_Close()                    ; close the connection to Outlook


General                                                         2011-05-18
--------------------------------------------------------------------------
* <will be added>