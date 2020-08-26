# Handling SharePoint Documents in Business Central

This repository contains a sample project that integrates Business Central with SharePoint.

This project is using "[Generic OAuth 2.0 Library](https://github.com/msnraju/BC-OAuth-2.0-Authorization)" as a dependency for [Microsoft Graph API](https://docs.microsoft.com/en-us/graph/use-the-api) Authentication Token.

## "Online Drives" ( page 50115 )
This page connects with the SharePoint root site and shows list of document libraries available in the root site.

![drives](/media/drives.png)

You can also navigate through folders and files by clicking on the drive name.

## "Online Drive Items" ( page 50116 )
This page connects with the SharePoint document library and displays list of folders and files in that library.

![drives](/media/drive-items.png)

You can perform the following operations: 
* Create a new folder
* Upload a file
* Download a file (click on the file name)
* Delete a folder / file

## Code in Action

![SharePoint Connect](/media/sharepoint-connect.gif)

For more information go to [msnJournals.com](https://www.msnjournals.com/post/how-to-connect-sharepoint-with-business-central)

