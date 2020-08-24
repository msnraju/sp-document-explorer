page 50115 "Online Drives"
{
    PageType = List;
    ApplicationArea = All;
    UsageCategory = Administration;
    SourceTable = "Online Drive";
    SourceTableTemporary = true;
    Editable = false;

    layout
    {
        area(content)
        {
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
                field("Drive Type"; Rec.driveType)
                {
                    Caption = 'Type';
                    ApplicationArea = All;
                }
                field("Last Modified DateTime"; Rec.lastModifiedDateTime)
                {
                    Caption = 'Last Modified DateTime';
                    ApplicationArea = All;
                }
            }
        }
    }

    trigger OnOpenPage()
    var
        TempOnlineDrive: Record "Online Drive" temporary;
        OnlineDriveAPI: Codeunit "Online Drive API";
    begin
        AccessToken := OnlineDriveAPI.GetAccessToken('AZUREAD');
        if OnlineDriveAPI.FetchDrives(AccessToken, TempOnlineDrive) then
            Rec.Copy(TempOnlineDrive, true);
    end;

    var
        AccessToken: Text;

    local procedure OpenDriveItems()
    var
        OnlineDriveItems: Page "Online Drive Items";
    begin
        OnlineDriveItems.SetProperties(AccessToken, '', Id, '');
        OnlineDriveItems.Run();
    end;
}