# ConvertOldOfficeFiles
Convert Excel and Word files in old Office 2003 format automatically to the new OpenXML format

This is my first github project, so please bear with me...

Sometime ago we decided to block mail attachments with old office formats like *.doc and *.xls in our company as they could contain macros and therefore could be a high risk for getting unwanted malware. Unfortunately our own users also sent out documents in these formats as they're really comfortable in using old files, copy them, change and send it to their contacts. The only way to stop this bevavior was to convert all files in these formats on our file server. Having tons of files this could be a very nasty task and therefore I decided to develop this small tool.

The tool let you scan a given directory and all its subdirectories for files like *.doc and *.xls. Matching files can be converted by opening a Word or Excel COM instance in the background and saving the files in the new OpenXML format. Files containing macros were saved in the appropiate format (docm or xlsm). Files with a matching extension containing non-Office format (like text files with an extension .doc) will be skipped.

I hope this tiny tool will be useful for anybody else in a similar situation.
