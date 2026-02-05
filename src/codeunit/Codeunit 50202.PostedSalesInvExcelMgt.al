codeunit 50202 "Posted Sales Inv Excel Mgt TCL"
{


    procedure ExportPostedSalesInvoices(SalesInvHeader: Record "Sales Invoice Header")
    
    
    begin
        InitWorkbook();
        WriteHeaderSheetHeaders();

        if SalesInvHeader.FindSet() then
            repeat
                WriteHeaderRow(SalesInvHeader);
                WriteLineRows(SalesInvHeader);
            until SalesInvHeader.Next() = 0;

        FinishWorkbook();
    end;

    local procedure InitWorkbook()
    begin
        ExcelBuf.DeleteAll();
        ExcelBuf.CreateNewBook('Posted Sales Invoices Export');
        ExcelBuf.SelectOrAddSheet('Headers');
    end;

    local procedure FinishWorkbook()
    begin
        ExcelBuf.WriteSheet('Headers', CompanyName, UserId);
        ExcelBuf.CloseBook();
        ExcelBuf.OpenExcel();
    end;

    local procedure WriteHeaderSheetHeaders()
    begin
        ExcelBuf.NewRow();
        ExcelBuf.AddColumn('Order No.', false, '', true, false, false, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Sell-to Customer', false, '', true, false, false, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Posting Date', false, '', true, false, false, '', ExcelBuf."Cell Type"::Date);
    end;

    local procedure WriteHeaderRow(SalesInvHeader: Record "Sales Invoice Header")
    begin
        ExcelBuf.SelectOrAddSheet('Headers');
        ExcelBuf.NewRow();
        ExcelBuf.AddColumn(SalesInvHeader."Order No.", false, '', false, false, false, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(SalesInvHeader."Sell-to Customer Name", false, '', false, false, false, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn(SalesInvHeader."Posting Date", false, '', false, false, false, '', ExcelBuf."Cell Type"::Date);
    end;

    local procedure WriteLineRows(SalesInvHeader: Record "Sales Invoice Header")
    var
        SalesInvLine: Record "Sales Invoice Line";
    begin
        ExcelBuf.SelectOrAddSheet('Lines');

        if ExcelBuf.GetRowCount() = 0 then
            WriteLineSheetHeaders();

        SalesInvLine.SetRange("Document No.", SalesInvHeader."No.");
        if SalesInvLine.FindSet() then
            repeat
                ExcelBuf.NewRow();
                ExcelBuf.AddColumn(SalesInvLine."Document No.", false, '', false, false, false, '', ExcelBuf."Cell Type"::Text);
                ExcelBuf.AddColumn(Format(SalesInvLine."Entry Type"), false, '', false, false, false, '', ExcelBuf."Cell Type"::Text);
                ExcelBuf.AddColumn(SalesInvLine."No.", false, '', false, false, false, '', ExcelBuf."Cell Type"::Text);
                ExcelBuf.AddColumn(SalesInvLine.Description, false, '', false, false, false, '', ExcelBuf."Cell Type"::Text);
                ExcelBuf.AddColumn(SalesInvLine.Quantity, false, '', false, false, false, '', ExcelBuf."Cell Type"::Number);
                ExcelBuf.AddColumn(SalesInvLine.Amount, false, '', false, false, false, '', ExcelBuf."Cell Type"::Number);
            until SalesInvLine.Next() = 0;
    end;

    local procedure WriteLineSheetHeaders()
    begin
        ExcelBuf.NewRow();
        ExcelBuf.AddColumn('Document No.', false, '', true, false, false, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Entry Type', false, '', true, false, false, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('No.', false, '', true, false, false, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Description', false, '', true, false, false, '', ExcelBuf."Cell Type"::Text);
        ExcelBuf.AddColumn('Quantity', false, '', true, false, false, '', ExcelBuf."Cell Type"::Number);
        ExcelBuf.AddColumn('Amount', false, '', true, false, false, '', ExcelBuf."Cell Type"::Number);
    end;

    var
        ExcelBuf: Record "Excel Buffer" temporary;
}