// Fees – one row per fee charged per semester
let
    Source = Excel.Workbook(File.Contents(DataFile), null, true),
    Fees_Sheet = Source{[Item = "Fees", Kind = "Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(Fees_Sheet, [PromoteAllScalars = true]),
    // Date, not Date/Time – otherwise the DueDate comparisons in DAX drift
    #"Changed Type" = Table.TransformColumnTypes(
        #"Promoted Headers",
        {{"FeeID", type text}, {"StudentID", type text}, {"Semester", type text},
         {"FeeType", type text}, {"Amount", Int64.Type}, {"DueDate", type date}}
    )
in
    #"Changed Type"
