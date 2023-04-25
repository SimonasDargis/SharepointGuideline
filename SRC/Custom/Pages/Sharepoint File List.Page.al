page 50004 "Sharepoint File List"
{
    PageType = List;
    SourceTable = "Sharepoint File List";
    InsertAllowed = false;
    DeleteAllowed = false;
    ModifyAllowed = false;
    SourceTableTemporary = true;
    ApplicationArea = All;
    UsageCategory = Lists;

    layout
    {
        area(Content)
        {
            repeater(GroupName)
            {
                field(Folder; Rec.Folder)
                {
                    ApplicationArea = All;
                }
                field(Name; Rec.Title)
                {
                    ApplicationArea = All;

                    trigger OnDrillDown()
                    var
                        FileList: Page "Sharepoint File List";
                    begin
                        if Rec.Folder then begin
                            FileList.LookupMode(true);
                            FileList.SetParentFolderURL(Rec."Server Relative Url");
                            FileList.RunModal();
                        end;
                    end;
                }
                field(OdataId; Rec.OdataId)
                {
                    ApplicationArea = All;
                }
                field("Server Relative Url"; Rec."Server Relative Url")
                {
                    ApplicationArea = All;
                }
            }
        }
    }

    actions
    {
        area(Processing)
        {
            action(OpenFile)
            {
                ApplicationArea = All;
                Caption = 'Open File';
                trigger OnAction()
                var
                    SharepointMgt: Codeunit "Sharepoint Management";
                begin
                    SharepointMgt.OpenFile(Rec."Server Relative Url");
                end;
            }
            action(CreateFile)
            {
                ApplicationArea = All;
                Caption = 'Create File';

                trigger OnAction()
                var
                    SharePointFile: Record "SharePoint File" temporary;
                    IS: InStream;
                    OS: OutStream;
                    TempBlob: Codeunit "Temp Blob";
                    SharepointMgt: Codeunit "Sharepoint Management";
                begin
                    OS := TempBlob.CreateOutStream();
                    OS.Write('Testing file contents');
                    IS := TempBlob.CreateInStream();

                    if SharepointMgt.SaveFile(ParentFolderURL, 'New File.txt', IS) then
                        Message('File created successfully!');
                end;
            }
        }
    }

    trigger OnOpenPage()
    var
        SharepointFile: Record "SharePoint File" temporary;
        SharepointFolder: Record "SharePoint Folder" temporary;
        SharepointMgt: Codeunit "Sharepoint Management";
    begin
        if ParentFolderURL <> '' then begin
            SharepointMgt.GetFilesFromServerRelativeURL(ParentFolderURL, SharepointFolder, SharepointFile);
        end else begin
            ParentFolderURL := SharepointMgt.GetDocumentsRootFiles(SharepointFolder, SharepointFile);
        end;

        if SharepointFolder.FindSet() then begin
            repeat
                Rec.Init();
                Rec.Id := SharepointFolder."Unique Id";
                Rec.Title := SharepointFolder.Name;
                Rec.OdataId := SharepointFolder.OdataId;
                Rec."Server Relative Url" := SharepointFolder."Server Relative Url";
                Rec.Folder := true;
                Rec.Insert();
            until SharepointFolder.Next() = 0;
        end;

        if SharepointFile.FindSet() then begin
            repeat
                Rec.Init();
                Rec.Id := SharepointFile."Unique Id";
                Rec.Title := SharepointFile.Name;
                Rec.OdataId := SharepointFile.OdataId;
                Rec."Server Relative Url" := SharepointFile."Server Relative Url";
                Rec.Insert();
            until SharepointFile.Next() = 0;
        end;
    end;

    procedure SetParentFolderURL(NewURL: Text)
    begin
        ParentFolderURL := NewURL;
    end;

    var
        ParentFolderURL: Text;

}