page 92300 "Upload ZIP with documents"
{
    Caption = 'SharePoint Upload ZIP with Documents (Custom)';
    UsageCategory = Administration;
    ApplicationArea = all;
    PageType = Card;
    layout
    {
        area(Content)
        {
            field(SourceTable; SourceTable)
            {
                Caption = 'Upload Into Table';
                ApplicationArea = all;
            }
            field(IdentificationField; IdentificationField)
            {
                Caption = 'Identification Field No.';
                ApplicationArea = all;
            }
            field(Instructions; Instructions)
            {
                ShowCaption = false;
                MultiLine = true;
            }
        }
    }
    actions
    {
        area(Processing)
        {
            action(Upload)
            {
                Caption = 'Upload ZIP';
                ApplicationArea = all;
                Promoted = true;
                PromotedCategory = Process;
                PromotedOnly = true;
                trigger OnAction()
                begin
                    if SourceTable = 0 then
                        error('Source Table is needed, cannot continue');
                    if IdentificationField = 0 then
                        error('Identification field is needed, cannot continue');
                    RunUpload();
                end;
            }
        }
    }
    procedure RunUpload()
    var
        InS: InStream;
        ZIP: Codeunit "Data Compression";
        EntryList: List of [Text];
        Entry: Text;
        Parts: List of [text];
        Part: Text;
        i: Integer;
        Ref: RecordRef;
        FR: FieldRef;
        OutS: OutStream;
        SP: Codeunit "SharePoint EFQ";
        TempBlob: Codeunit "Temp Blob";
        InS2: InStream;
        Folder: Text;
        Mapping: Record "Table Mapping EFQ";
        Done: Boolean;
        Window: Dialog;
        FileName: Text;
        Token: Text;
    begin
        SP.GetAccessTokenAgain(Token);
        SP.StoreAccessToken(Token);
        Window.Open('Importing #1##############################');
        if UploadIntoStream('Select Zip File', '', '', FileName, InS) then begin
            ZIP.OpenZipArchive(InS, false);
            ZIP.GetEntryList(EntryList);
            foreach Entry in EntryList do
                if not Entry.EndsWith('/') then begin
                    //message('%1', Entry);
                    Parts := Entry.Split('/');
                    Done := false;
                    for i := Parts.Count downto 1 do
                        if not Done then begin
                            Part := Parts.Get(i);
                            Ref.Open(SourceTable);
                            FR := Ref.Field(IdentificationField);
                            FR.SetFilter(Part);
                            if Ref.FindFirst() then begin
                                Window.Update(1, Parts.get(Parts.Count));
                                clear(TempBlob);
                                TempBlob.CreateOutStream(OutS);
                                ZIP.ExtractEntry(Entry, OutS);
                                TempBlob.CreateInStream(InS);

                                Folder := SP.GetFolderForRecord(Ref, true);
                                SP.GetTableMapping(Mapping, Ref);
                                if not SP.UploadFile(Mapping, Folder, InS, SP.SanitizeNameNoDot(Parts.Get(Parts.Count)), false) then
                                    error(UploadFailedErr)
                                else
                                    SP.FilloutCustomColumns(Mapping, Folder, SP.SanitizeNameNoDot(Parts.Get(Parts.Count)), Ref);
                                Done := true;
                            end;
                            REf.Close();
                        end;
                end;
        end;
    end;

    var
        SourceTable: Integer;
        IdentificationField: Integer;
        UploadFailedErr: Label 'Upload to SharePoint failed';
        Instructions: Label 'This function take a ZIP file with files organized in subfolders. The upload will try to find a record in the specified table with a lookup based on the folders. So if a file is located in a folder called "10000" then this function will upload that file into Customer 10000.';

}