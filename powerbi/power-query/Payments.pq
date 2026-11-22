// Payments – one row per payment received
let
    Source = Excel.Workbook(File.Contents(DataFile), null, true),
    Payments_Sheet = Source{[Item = "Payments", Kind = "Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(Payments_Sheet, [PromoteAllScalars = true]),
    // 9 messy spellings → Bank Deposit, Cash, Online
    #"Trimmed Text" = Table.TransformColumns(#"Promoted Headers", {{"Method", Text.Trim, type text}}),
    #"Capitalized Each Word" = Table.TransformColumns(#"Trimmed Text", {{"Method", Text.Proper, type text}}),
    // Blank ReceiptNo values are kept on purpose: they are an exception to report
    #"Changed Type" = Table.TransformColumnTypes(
        #"Capitalized Each Word",
        {{"PaymentID", type text}, {"FeeID", type text}, {"PaymentDate", type date},
         {"AmountPaid", Int64.Type}, {"ReceiptNo", type text}, {"RecordedBy", type text}}
    )
in
    #"Changed Type"
