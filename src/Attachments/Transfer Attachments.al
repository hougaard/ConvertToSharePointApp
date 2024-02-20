report 92300 "Transfer Attachments"
{
    ProcessingOnly = true;
    Caption = 'Transfer Attachments to SharePoint (Custom)';
    UsageCategory = Administration;
    ApplicationArea = all;
    dataset
    {
        dataitem("Document Attachment"; "Document Attachment")
        {
            RequestFilterFields = "Table ID", "No.", "Attached Date", "Document Type", User;
            trigger OnPreDataItem()
            begin
                Window.Open(ImportingLbl);
            end;

            trigger OnAfterGetRecord()
            begin
                begin
                    Ref.Open("Document Attachment"."Table ID");
                    KR := Ref.KeyIndex(1);
                    case KR.FieldCount() of
                        1:
                            begin
                                FR := KR.FieldIndex(1);
                                FR.Value := "Document Attachment"."No.";
                            end;
                        2:
                            if "Document Attachment"."Table ID" in [36, 38] then begin
                                FR := KR.FieldIndex(1);
                                if FR.Type = FieldType::Option then
                                    FR.Value := "Document Attachment"."Document Type";
                                FR := KR.FieldIndex(2);
                                FR.Value := "Document Attachment"."No.";
                            end else
                                error(WrongTable2Err);
                        else
                            if KR.FieldCount() <> 1 then
                                error(WrongTableErr);
                    end;
                    if Ref.Find('=') then begin
                        Window.Update(1, "Document Attachment"."File Name");
                        clear(TempBlob);
                        Folder := SP.GetFolderForRecord(Ref, true);
                        TempBlob.CreateOutStream(DocumentStream);
                        "Document Attachment"."Document Reference ID".ExportStream(DocumentStream);
                        TempBlob.CreateInStream(InS);
                        SP.GetTableMapping(Mapping, Ref);
                        if not SP.UploadFile(Mapping, Folder, InS, sp.SanitizeName("Document Attachment"."File Name") + '.' + "Document Attachment"."File Extension", false) then
                            error(UploadFailedErr)
                        else
                            SP.FilloutCustomColumns(Mapping, Folder, SP.SanitizeName("Document Attachment"."File Name" + '.' + "Document Attachment"."File Extension"), Ref);
                    end;
                    Ref.Close();
                end;
            end;
        }
    }
    var
        Mapping: Record "Table Mapping EFQ";
        TempBlob: Codeunit "Temp Blob";
        SP: Codeunit "SharePoint EFQ";
        Ref: RecordRef;
        FR: FieldRef;
        KR: KeyRef;
        UploadFailedErr: Label 'Upload to SharePoint failed';
        WrongTableErr: Label 'Only tables with 1-field primary key can be used for importing attachments, aborting!';
        WrongTable2Err: Label 'Only sales header and purchase header tables with 2-field primary keys can be used for importing attachments, aborting!';
        ImportingLbl: Label 'Importing attachments to SharePoint #1###############################', Comment = '#1# is filename';
        Folder: Text;
        InS: InStream;
        DocumentStream: OutStream;
        Window: Dialog;
}