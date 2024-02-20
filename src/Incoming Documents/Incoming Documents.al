report 92301 "Tranfer Incoming Documents"
{
    ProcessingOnly = true;
    Caption = 'Transfer Incoming Documents to SharePoint (Custom)';
    UsageCategory = Administration;
    ApplicationArea = all;
    dataset
    {
        dataitem(IncomingDocument; "Incoming Document")
        {
            RequestFilterFields = "Document Date", "Document No.", "Document Type", "Vendor No.", "Entry No.", "Order No.";
            trigger OnPreDataItem()
            begin
                Window.Open(ImportingLbl);
            end;

            trigger OnAfterGetRecord()
            var
                Attachment: Record "Incoming Document Attachment";
            begin
                if Mapping.Get(IncomingDocument."Related Record ID".TableNo) then begin
                    Attachment.Setrange("Incoming Document Entry No.", IncomingDocument."Entry No.");
                    if Attachment.FindSet() then
                        repeat
                            Ref.Open(IncomingDocument."Related Record ID".TableNo);
                            if Ref.Get(IncomingDocument."Related Record ID") then begin
                                Window.Update(1, Attachment.Name);
                                clear(TempBlob);
                                Folder := SP.GetFolderForRecord(Ref, true);
                                Attachment.CalcFields(Content);
                                Attachment.Content.CreateInStream(InS);
                                SP.GetTableMapping(Mapping, Ref);
                                if not SP.UploadFile(Mapping, Folder, InS, SP.SanitizeName(Attachment.name) + '.' + Attachment."File Extension", false) then
                                    UploadErrorCounter += 1
                                else begin
                                    UploadCounter += 1;
                                    SP.FilloutCustomColumns(Mapping, Folder, SP.SanitizeName(Attachment.name) + '.' + Attachment."File Extension", Ref);
                                end;
                            end;
                            Ref.Close();
                        until Attachment.Next() = 0;
                end;
            end;
        }
    }
    trigger OnPreReport()
    begin
        SP.GetAccessTokenAgain(Token);
        Sp.StoreAccessToken(Token);
    end;

    var
        Token: Text;
        Mapping: Record "Table Mapping EFQ";
        TempBlob: Codeunit "Temp Blob";
        SP: Codeunit "SharePoint EFQ";
        Ref: RecordRef;
        FR: FieldRef;
        KR: KeyRef;
        UploadErrorCounter: Integer;
        UploadCounter: Integer;
        UploadFailedErr: Label 'Upload to SharePoint failed';
        ImportingLbl: Label 'Importing incoming documents to SharePoint #1###############################', Comment = '#1# is filename';
        WrongTableErr: Label 'Only tables with 1-field primary key can be used for importing attachments, aborting!';
        WrongTable2Err: Label 'Only sales header and purchase header tables with 2-field primary keys can be used for importing attachments, aborting!';
        Folder: Text;
        InS: InStream;
        DocumentStream: OutStream;
        Window: Dialog;
}