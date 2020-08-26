codeunit 50115 "Online Drive API"
{
    var
        DrivesUrl: Label 'https://graph.microsoft.com/v1.0/drives', Locked = true;
        DrivesItemsUrl: Label 'https://graph.microsoft.com/v1.0/drives/%1/root/children', Comment = '%1 = Drive ID', Locked = true;
        DrivesChildItemsUrl: Label 'https://graph.microsoft.com/v1.0/drives/%1/items/%2/children', Comment = '%1 = Drive ID, %2 = Item ID', Locked = true;
        UploadUrl: Label 'https://graph.microsoft.com/v1.0/drives/%1/items/root:/%2:/content', Comment = '%1 = Drive ID, %2 = File Name', Locked = true;
        DownloadUrl: Label 'https://graph.microsoft.com/v1.0/drives/%1/items/%2/content', Comment = '%1 = Drive ID, %2 = Item ID', Locked = true;
        DeleteUrl: Label 'https://graph.microsoft.com/v1.0/drives/%1/items/%2', Comment = '%1 = Drive ID, %2 = Item ID', Locked = true;
        CreateFolderUrl: Label 'https://graph.microsoft.com/v1.0/drives/%1/items/%2/children', Comment = '%1 = Drive ID, %2 = Item ID', Locked = true;
        CreateRootFolderUrl: Label 'https://graph.microsoft.com/v1.0/drives/%1/root/children', Comment = '%1 = Drive ID', Locked = true;

    procedure GetAccessToken(AppCode: Code[20]): Text
    var
        OAuth20Application: Record "OAuth 2.0 Application";
        OAuth20AppHelper: Codeunit "OAuth 2.0 App. Helper";
        MessageText: Text;
    begin
        OAuth20Application.Get(AppCode);
        if not OAuth20AppHelper.RequestAccessToken(OAuth20Application, MessageText) then
            Error(MessageText);

        exit(OAuth20AppHelper.GetAccessToken(OAuth20Application));
    end;

    procedure UploadFile(
        AccessToken: Text;
        DriveID: Text;
        ParentID: Text;
        FolderPath: Text;
        FileName: Text;
        var Stream: InStream;
        var OnlineDriveItem: Record "Online Drive Item"): Boolean
    var
        HttpClient: HttpClient;
        Headers: HttpHeaders;
        RequestMessage: HttpRequestMessage;
        RequestContent: HttpContent;
        ResponseMessage: HttpResponseMessage;
        JsonResponse: JsonObject;
        IsSucces: Boolean;
        ResponseText: Text;
    begin
        Headers := HttpClient.DefaultRequestHeaders();
        Headers.Add('Authorization', StrSubstNo('Bearer %1', AccessToken));

        RequestMessage.SetRequestUri(
            StrSubstNo(
                UploadUrl,
                DriveID,
                StrSubstNo('%1/%2', FolderPath, FileName)));
        RequestMessage.Method := 'PUT';

        RequestContent.WriteFrom(Stream);
        RequestMessage.Content := RequestContent;

        if HttpClient.Send(RequestMessage, ResponseMessage) then
            if ResponseMessage.IsSuccessStatusCode() then begin
                if ResponseMessage.Content.ReadAs(ResponseText) then begin
                    IsSucces := true;
                    if JsonResponse.ReadFrom(ResponseText) then
                        ReadDriveItem(JsonResponse, DriveID, ParentID, OnlineDriveItem);
                end;
            end else
                if ResponseMessage.Content.ReadAs(ResponseText) then
                    JsonResponse.ReadFrom(ResponseText);

        exit(IsSucces);
    end;

    procedure DownloadFile(AccessToken: Text; DriveID: Text; ItemID: Text; var Stream: InStream): Boolean
    var
        TempBlob: Codeunit "Temp Blob";
        OStream: OutStream;
        JsonResponse: JsonObject;
        Content: Text;
        NewDownloadUrl: Text;
    begin
        NewDownloadUrl := StrSubstNo(DownloadUrl, DriveID, ItemID);
        if GetResponse(AccessToken, NewDownloadUrl, Stream) then
            exit(true);
    end;

    procedure CreateDriveFolder(
        AccessToken: Text;
        DriveID: Text;
        ItemID: Text;
        FolderName: Text;
        var OnlineDriveItem: Record "Online Drive Item"): Boolean
    var
        HttpClient: HttpClient;
        Headers: HttpHeaders;
        ContentHeaders: HttpHeaders;
        RequestMessage: HttpRequestMessage;
        RequestContent: HttpContent;
        ResponseMessage: HttpResponseMessage;
        ResponseText: Text;
        JsonBody: JsonObject;
        RequestText: Text;
        EmptyObject: JsonObject;
        JsonResponse: JsonObject;
    begin
        Headers := HttpClient.DefaultRequestHeaders();
        Headers.Add('Authorization', StrSubstNo('Bearer %1', AccessToken));
        if ItemID = '' then
            RequestMessage.SetRequestUri(StrSubstNo(CreateRootFolderUrl, DriveID))
        else
            RequestMessage.SetRequestUri(StrSubstNo(CreateFolderUrl, DriveID, ItemID));
        RequestMessage.Method := 'POST';

        // Body
        JsonBody.Add('name', FolderName);
        JsonBody.Add('folder', EmptyObject);
        JsonBody.WriteTo(RequestText);
        RequestContent.WriteFrom(RequestText);

        // Content Headers
        RequestContent.GetHeaders(ContentHeaders);
        ContentHeaders.Remove('Content-Type');
        ContentHeaders.Add('Content-Type', 'application/json');

        RequestMessage.Content := RequestContent;

        if HttpClient.Send(RequestMessage, ResponseMessage) then
            if ResponseMessage.IsSuccessStatusCode() then begin
                if ResponseMessage.Content.ReadAs(ResponseText) then begin
                    if JsonResponse.ReadFrom(ResponseText) then
                        ReadDriveItem(JsonResponse, DriveID, ItemID, OnlineDriveItem);

                    exit(true);
                end;
            end;
    end;

    procedure DeleteDriveItem(AccessToken: Text; DriveID: Text; ItemID: Text): Boolean
    var
        HttpClient: HttpClient;
        Headers: HttpHeaders;
        RequestMessage: HttpRequestMessage;
        ResponseMessage: HttpResponseMessage;
    begin
        Headers := HttpClient.DefaultRequestHeaders();
        Headers.Add('Authorization', StrSubstNo('Bearer %1', AccessToken));

        RequestMessage.SetRequestUri(StrSubstNo(DeleteUrl, DriveID, ItemID));
        RequestMessage.Method := 'DELETE';


        if HttpClient.Send(RequestMessage, ResponseMessage) then
            if ResponseMessage.IsSuccessStatusCode() then
                exit(true);
    end;

    procedure FetchDrives(AccessToken: Text; var Drive: Record "Online Drive"): Boolean
    var
        JsonResponse: JsonObject;
        JToken: JsonToken;
    begin
        if HttpGet(AccessToken, DrivesUrl, JsonResponse) then begin
            if JsonResponse.Get('value', JToken) then
                ReadDrives(JToken.AsArray(), Drive);

            exit(true);
        end;
    end;

    procedure FetchDrivesItems(AccessToken: Text; DriveID: Text; var DriveItem: Record "Online Drive Item"): Boolean
    var
        JsonResponse: JsonObject;
        JToken: JsonToken;
        IsSucces: Boolean;
    begin
        if HttpGet(AccessToken, StrSubstNo(DrivesItemsUrl, DriveID), JsonResponse) then begin
            if JsonResponse.Get('value', JToken) then
                ReadDriveItems(JToken.AsArray(), DriveID, '', DriveItem);

            exit(true);
        end;
    end;

    procedure FetchDrivesChildItems(
        AccessToken: Text;
        DriveID: Text;
        ItemID: Text;
        var DriveItem: Record "Online Drive Item"): Boolean
    var
        JsonResponse: JsonObject;
        JToken: JsonToken;
        IsSucces: Boolean;
    begin
        if HttpGet(AccessToken, StrSubstNo(DrivesChildItemsUrl, DriveID, ItemID), JsonResponse) then begin
            if JsonResponse.Get('value', JToken) then
                ReadDriveItems(JToken.AsArray(), DriveID, ItemID, DriveItem);

            exit(true);
        end;
    end;

    local procedure GetDriveID(var Drive: Record "Online Drive"; Name: Text): Text
    begin
        Drive.SetRange(Name, Name);
        if Drive.FindFirst() then
            exit(Drive.Id);
    end;

    local procedure GetItemID(var DriveItem: Record "Online Drive Item"; DriveID: Text; Name: Text): Text
    begin
        exit(GetItemID(DriveItem, DriveID, '', Name));
    end;

    local procedure GetItemID(
        var DriveItem: Record "Online Drive Item";
        DriveID: Text;
        ItemID: Text;
        Name: Text): Text
    begin
        DriveItem.SetRange(driveID, DriveID);
        DriveItem.SetRange(parentId, ItemID);
        DriveItem.SetRange(Name, Name);
        if DriveItem.FindFirst() then
            exit(DriveItem.Id);
    end;

    local procedure ReadDrives(JDrives: JsonArray; var Drive: Record "Online Drive")
    var
        JDriveItem: JsonToken;
        JDrive: JsonObject;
        JToken: JsonToken;
    begin
        foreach JDriveItem in JDrives do begin
            JDrive := JDriveItem.AsObject();

            Drive.Init();
            if JDrive.Get('id', JToken) then
                Drive.Id := JToken.AsValue().AsText();
            if JDrive.Get('name', JToken) then
                Drive.Name := JToken.AsValue().AsText();
            if JDrive.Get('description', JToken) then
                Drive.description := JToken.AsValue().AsText();
            if JDrive.Get('driveType', JToken) then
                Drive.driveType := JToken.AsValue().AsText();
            if JDrive.Get('createdDateTime', JToken) then
                Drive.createdDateTime := JToken.AsValue().AsDateTime();
            if JDrive.Get('lastModifiedDateTime', JToken) then
                Drive.lastModifiedDateTime := JToken.AsValue().AsDateTime();
            if JDrive.Get('webUrl', JToken) then
                Drive.webUrl := JToken.AsValue().AsText();
            Drive.Insert();
        end;
    end;

    local procedure ReadDriveItems(
        JDriveItems: JsonArray;
        DriveID: Text;
        ParentID: Text;
        var DriveItem: Record "Online Drive Item")
    var
        JToken: JsonToken;
    begin
        foreach JToken in JDriveItems do
            ReadDriveItem(JToken.AsObject(), DriveID, ParentID, DriveItem);
    end;

    local procedure ReadDriveItem(
        JDriveItem: JsonObject;
        DriveID: Text;
        ParentID: Text;
        var DriveItem: Record "Online Drive Item")
    var
        JFile: JsonObject;
        JToken: JsonToken;
    begin

        DriveItem.Init();
        DriveItem.driveId := DriveID;
        DriveItem.parentId := ParentID;

        if JDriveItem.Get('id', JToken) then
            DriveItem.Id := JToken.AsValue().AsText();
        if JDriveItem.Get('name', JToken) then
            DriveItem.Name := JToken.AsValue().AsText();

        if JDriveItem.Get('size', JToken) then
            DriveItem.size := JToken.AsValue().AsBigInteger();

        if JDriveItem.Get('file', JToken) then begin
            DriveItem.IsFile := true;
            JFile := JToken.AsObject();
            if JFile.Get('mimeType', JToken) then
                DriveItem.mimeType := JToken.AsValue().AsText();
        end;

        if JDriveItem.Get('createdDateTime', JToken) then
            DriveItem.createdDateTime := JToken.AsValue().AsDateTime();
        if JDriveItem.Get('webUrl', JToken) then
            DriveItem.webUrl := JToken.AsValue().AsText();
        DriveItem.Insert();
    end;

    local procedure GetResponse(AccessToken: Text; Url: Text; var Stream: InStream): Boolean
    var
        Client: HttpClient;
        Headers: HttpHeaders;
        RequestMessage: HttpRequestMessage;
        ResponseMessage: HttpResponseMessage;
        RequestContent: HttpContent;
        IsSucces: Boolean;
    begin
        Headers := Client.DefaultRequestHeaders();
        Headers.Add('Authorization', StrSubstNo('Bearer %1', AccessToken));

        RequestMessage.SetRequestUri(Url);
        RequestMessage.Method := 'GET';

        if Client.Send(RequestMessage, ResponseMessage) then
            if ResponseMessage.IsSuccessStatusCode() then begin
                if ResponseMessage.Content.ReadAs(Stream) then
                    IsSucces := true;
            end else
                ResponseMessage.Content.ReadAs(Stream);

        exit(IsSucces);
    end;

    local procedure HttpGet(AccessToken: Text; Url: Text; var JResponse: JsonObject): Boolean
    var
        Client: HttpClient;
        Headers: HttpHeaders;
        RequestMessage: HttpRequestMessage;
        ResponseMessage: HttpResponseMessage;
        RequestContent: HttpContent;
        ResponseText: Text;
        IsSucces: Boolean;
    begin
        Headers := Client.DefaultRequestHeaders();
        Headers.Add('Authorization', StrSubstNo('Bearer %1', AccessToken));

        RequestMessage.SetRequestUri(Url);
        RequestMessage.Method := 'GET';

        if Client.Send(RequestMessage, ResponseMessage) then
            if ResponseMessage.IsSuccessStatusCode() then begin
                if ResponseMessage.Content.ReadAs(ResponseText) then
                    IsSucces := true;
            end else
                ResponseMessage.Content.ReadAs(ResponseText);

        JResponse.ReadFrom(ResponseText);
        exit(IsSucces);
    end;
}