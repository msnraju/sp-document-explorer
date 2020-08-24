page 50116 "Online Drive Items"
{
    PageType = List;
    ApplicationArea = All;
    UsageCategory = Administration;
    SourceTable = "Online Drive Item";
    SourceTableTemporary = true;
    Editable = false;

    layout
    {
        area(content)
        {
            group(Title)
            {
                ShowCaption = false;
                Visible = FolderPath <> '';

                field(Parent; FolderPath)
                {
                    ShowCaption = false;
                }
            }

            repeater(Control1)
            {
                ShowCaption = false;

                field("Name"; Rec.Name)
                {
                    ApplicationArea = All;
                    Caption = 'Name';

                    trigger OnDrillDown()
                    begin
                        OpenDriveItems();
                    end;
                }
                field("Is File"; Rec.isFile)
                {
                    Caption = 'Is File';
                    ApplicationArea = All;
                }
                field("Created DateTime"; Rec.createdDateTime)
                {
                    Caption = 'Created DateTime';
                    ApplicationArea = All;
                }
            }
        }
    }

    actions
    {
        area(Processing)
        {
            action("Create a new folder")
            {
                ApplicationArea = All;

                trigger OnAction()
                var
                    OnlineCreateDriveFolder: Page "Online Create Drive Folder";
                    FolderName: Text;
                begin
                    if OnlineCreateDriveFolder.RunModal() = Action::OK then begin
                        FolderName := OnlineCreateDriveFolder.GetFolderName();
                        if FolderName = '' then
                            exit;

                        OnlineDriveAPI.CreateDriveFolder(AccessToken, driveId, parentId, FolderName, Rec);
                    end;

                end;
            }
            action("Delete the selected item")
            {
                ApplicationArea = All;
                ShortcutKey = 'shift + f4';

                trigger OnAction()
                begin
                    if Confirm(ConfirmMsg, false) then
                        if OnlineDriveAPI.DeleteDriveItem(AccessToken, driveId, id) then
                            Delete();
                end;
            }
            action("Upload a File")
            {
                ApplicationArea = All;

                trigger OnAction()
                var
                    FromFile: Text;
                    Stream: InStream;
                begin
                    if UploadIntoStream('Select a File', '', '', FromFile, Stream) then
                        OnlineDriveAPI.UploadFile(AccessToken, driveId, parentId, FolderPath, FromFile, Stream, Rec);
                end;
            }
        }
    }

    var
        OnlineDriveAPI: Codeunit "Online Drive API";
        AccessToken: Text;
        FolderPath: Text;
        ConfirmMsg: Label 'Do you want to delete the selected item ?';

    procedure SetProperties(NewAccessToken: Text; NewFolderPath: Text; DriveID: Text; ParentID: Text)
    var
        TempOnlineDriveItem: Record "Online Drive Item" temporary;
    begin
        AccessToken := NewAccessToken;
        FolderPath := NewFolderPath;

        if ParentID = '' then
            OnlineDriveAPI.FetchDrivesItems(AccessToken, DriveID, TempOnlineDriveItem)
        else
            OnlineDriveAPI.FetchDrivesChildItems(AccessToken, DriveID, ParentID, TempOnlineDriveItem);

        Rec.Copy(TempOnlineDriveItem, true);
    end;

    local procedure OpenDriveItems()
    var
        OnlineDriveItems: Page "Online Drive Items";
        Stream: InStream;
    begin
        if not isFile then begin
            OnlineDriveItems.SetProperties(AccessToken, StrSubstNo('%1/%2', FolderPath, name), driveId, Id);
            OnlineDriveItems.Run();
        end else begin
            if OnlineDriveAPI.DownloadFile(AccessToken, driveId, id, Stream) then
                DownloadFromStream(Stream, '', '', '', name);
        end;
    end;
}