report 50200 "Export Posted Sales Invoices"
{
    UsageCategory = ReportsAndAnalysis;
    ApplicationArea = All;
    ProcessingOnly = true;

    dataset
    {
        dataitem(SalesInvHeader; "Sales Invoice Header")
        {
            RequestFilterFields = "Posting Date", "Sell-to Customer No.";

            trigger OnPreDataItem()
            begin
                CurrReport.Break(); // prevent default iteration 50100
            end;
        }
    }

    trigger OnPostReport()
    var
        SalesInvHeader: Record "Sales Invoice Header";
        ExcelMgt: Codeunit "Posted Sales Invoice Excel Mgt";
    begin
        // SalesInvHeader.CopyFilters(Rec);
        ExcelMgt.ExportPostedSalesInvoices(SalesInvHeader);
    end;
}