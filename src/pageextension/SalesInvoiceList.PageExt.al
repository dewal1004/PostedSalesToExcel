pageextension 50200 "Sales Invoice List TCL" extends "Posted Sales Invoices"  //143
{

    actions
    {
        addafter(Print)
        {
            action("Export to Excel TCL")
            {
                Caption = 'Export to Excel';
                Image = Export;
                ToolTip = 'Export the posted sales invoices to an Excel file.';
                ApplicationArea = All;

                trigger OnAction()
                var
                    PostedSalesInvoiceHandler: Codeunit "Posted Sales Invoice Excel Mgt";
                begin
                    PostedSalesInvoiceHandler.ExportPostedSalesInvoices(Rec);
                end;
            }

            action("ExportToExcel TCL")
            {
                ApplicationArea = All;
                Caption = 'Export to Excel';
                Image = "Export";

                trigger OnAction()
                var
                    SalesInvHeader: Record "Sales Invoice Header";
                begin
                    SalesInvHeader.SetRange("No.", Rec."No.");
                    Report.RunModal(Report::"Export Posted Sales Invoices", true, false, SalesInvHeader);
                end;
            }
        }
    }
}