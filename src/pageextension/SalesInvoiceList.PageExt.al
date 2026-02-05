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
        }
    }
}