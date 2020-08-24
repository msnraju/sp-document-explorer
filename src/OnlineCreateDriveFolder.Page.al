page 50117 "Online Create Drive Folder"
{
    PageType = StandardDialog;
    ApplicationArea = All;
    UsageCategory = Administration;

    layout
    {
        area(Content)
        {
            group(General)
            {
                ShowCaption = false;
                InstructionalText = 'Enter the folder name';

                field(Name; FolderName)
                {
                    ApplicationArea = All;
                }
            }
        }
    }

    var
        FolderName: Text;

    procedure GetFolderName(): Text
    begin
        exit(FolderName);
    end;
}