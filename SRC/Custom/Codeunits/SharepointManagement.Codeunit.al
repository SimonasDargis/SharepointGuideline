codeunit 50000 "Sharepoint Management"
{
    procedure SaveFile(FileDirectory: Text; FileName: Text; IS: InStream): Boolean
    var
        SharePointFile: Record "SharePoint File" temporary;
        IsSuccess: Boolean;
        Diag: Interface "HTTP Diagnostics";
    begin
        InitializeConnection(); //Initialize a connection if not connected
        IsSuccess := false; //Default status, is set to true if a file is save successfully.
        if SharePointClient.AddFileToFolder(FileDirectory, FileName, IS, SharePointFile) then //Use AddFileToFolder to create a file from a stream
            IsSuccess := true
        else begin
            Diag := SharePointClient.GetDiagnostics(); //Optional: used to get diagnostics, useful for debugging errors
            if (not Diag.IsSuccessStatusCode()) then //Check if the status is success, if not - failure to connect, display error message.
                Error(DiagError, Diag.GetErrorMessage());
        end;
    end;

    procedure OpenFile(FileDirectory: Text)
    var
        SharePointFolder: Record "SharePoint Folder" temporary;
        SharePointFile: Record "SharePoint File" temporary;
        FileMgt: Codeunit "File Management";
        FileName: Text;
        FileNotFoundErr: Label 'File not found.';
    begin
        InitializeConnection(); //Initialize a connection if not connected

        //Use the file management codeunit to seperate the file name from directory
        //Check if it's a file directoy and not a folder directory
        FileName := FileMgt.GetFileName(FileDirectory);
        if FileName = '' then
            Error(FileNotFoundErr);

        //The file name has been separated, change the full name to just the folder directory
        FileDirectory := FileMgt.GetDirectoryName(FileDirectory);

        //Check if possible to retrieve a list of files from a directory
        if not SharePointClient.GetFolderFilesByServerRelativeUrl(FileDirectory, SharePointFile) then
            Error(FileNotFoundErr);

        //Filter out the name we're looking for
        SharePointFile.SetRange(Name, FileName);
        if SharePointFile.FindFirst() then begin
            //Download the file if found
            SharePointClient.DownloadFileContent(SharePointFile.OdataId, FileName);
        end else
            Error(FileNotFoundErr);
    end;

    procedure GetDocumentsRootFiles(var SharepointFolder: Record "SharePoint Folder" temporary; var SharepointFile: Record "SharePoint File"): Text
    var
        SharePointList: Record "SharePoint List" temporary;
    begin
        InitializeConnection(); //Initialize a connection if not connected
        if SharePointClient.GetLists(SharePointList) then begin //Sharepoint List is empty, GetLists writes data
            SharePointList.SetRange(Title, 'Documents'); //We filter out the Documents list to access the documents library
            if SharePointList.FindFirst() then begin
                //Use GetDocumentLibraryRootFolder to get the root folder's server relative URL
                if SharePointClient.GetDocumentLibraryRootFolder(SharePointList.OdataId, SharePointFolder) then begin
                    //We then get the files from the root directory
                    //The SharepointFile record is filled with file records
                    //The SharePointFolder record is filled with folder records
                    SharePointClient.GetFolderFilesByServerRelativeUrl(SharePointFolder."Server Relative Url", SharePointFile);
                    SharePointClient.GetSubFoldersByServerRelativeUrl(SharePointFolder."Server Relative Url", SharePointFolder);
                    //You may loop through the records and create your own table to store both files and folders
                    exit(SharePointFolder."Server Relative Url"); //Exits with the root server relative URL
                end;
            end;
        end;
    end;

    procedure GetFilesFromServerRelativeURL(ServerRelativeURL: Text; var SharepointFolder: Record "SharePoint Folder" temporary;
        var SharepointFile: Record "SharePoint File")
    begin
        InitializeConnection(); //Initialize a connection if not connected
        //This function can be used if you already have a server relative URL
        //The URL can be retrieved with the previous function GetDocumentsRootFiles
        //This function requires only to use the same functions to just retrieve
        SharePointClient.GetFolderFilesByServerRelativeUrl(ServerRelativeURL, SharePointFile);
        SharePointClient.GetSubFoldersByServerRelativeUrl(ServerRelativeURL, SharePointFolder);
    end;

    local procedure InitializeConnection()
    var
        SharepointSetup: Record "Sharepoint Connector Setup";
        AadTenantId: Text;
        Diag: Interface "HTTP Diagnostics";
        SharePointList: Record "SharePoint List" temporary;
    begin
        if Connected then //A global variable Connected is used to store the value to prevent from needlessly repeating the function.
            exit;

        SharepointSetup.Get(); //Get Sharepoint Setup data

        AadTenantId := GetAadTenantNameFromBaseUrl(SharepointSetup."Sharepoint URL"); //Used to get an Azure Active Directory ID from a URL
        SharePointClient.Initialize(SharepointSetup."Sharepoint URL", GetSharePointAuthorization(AadTenantId)); //Initializes the client

        SharePointClient.GetLists(SharePointList); //We need to perform at least one action to get diagnostics data
                                                   //Otherwise, the GetDiagnostics function will just return 0
        Diag := SharePointClient.GetDiagnostics(); //Optional: used to get diagnostics, useful for debugging errors

        if (not Diag.IsSuccessStatusCode()) then //Check if the status is success, if not - failure to connect, display error message.
            Error(DiagError, Diag.GetErrorMessage());

        Connected := true; //Set the connection status to true so that we won't have to re-connect when already connected.
    end;

    local procedure GetSharePointAuthorization(AadTenantId: Text): Interface "SharePoint Authorization"
    var
        SharepointSetup: Record "Sharepoint Connector Setup";
        SharePointAuth: Codeunit "SharePoint Auth.";
        Scopes: List of [Text];
    begin
        SharepointSetup.Get(); //Get Sharepoint Setup data. Optionally, this can be made into a global variable as well.

        Scopes.Add('00000003-0000-0ff1-ce00-000000000000/.default'); //Using a default scope provided as an example
        //We return an authorization code that will be used to initialize the Sharepoint Client
        exit(SharePointAuth.CreateAuthorizationCode(AadTenantId, SharepointSetup."Client ID", SharepointSetup."Client Secret", Scopes));
    end;

    local procedure GetAadTenantNameFromBaseUrl(BaseUrl: Text): Text
    var
        Uri: Codeunit Uri;
        MySiteHostSuffixTxt: Label '-my.sharepoint.com', Locked = true;
        SharePointHostSuffixTxt: Label '.sharepoint.com', Locked = true;
        OnMicrosoftTxt: Label '.onmicrosoft.com', Locked = true;
        UrlInvalidErr: Label 'The Base Url %1 does not seem to be a valid SharePoint Online Url.', Comment = '%1=BaseUrl';
        Host: Text;
    begin
        //This procedure formats the sharepoint's site URL to a format accepted by CreateAuthorizationCode function
        // SharePoint Online format:  https://tenantname.sharepoint.com/SiteName/LibraryName/
        // SharePoint My Site format: https://tenantname-my.sharepoint.com/personal/user_name/
        Uri.Init(BaseUrl);
        Host := Uri.GetHost();
        if not Host.EndsWith(SharePointHostSuffixTxt) then
            Error(UrlInvalidErr, BaseUrl);
        if Host.EndsWith(MySiteHostSuffixTxt) then
            exit(CopyStr(Host, 1, StrPos(Host, MySiteHostSuffixTxt) - 1) + OnMicrosoftTxt);
        exit(CopyStr(Host, 1, StrPos(Host, SharePointHostSuffixTxt) - 1) + OnMicrosoftTxt);
    end;

    var
        Connected: Boolean;
        SharePointClient: Codeunit "SharePoint Client";
        DiagError: Label 'Sharepoint Management error:\\%1';
}