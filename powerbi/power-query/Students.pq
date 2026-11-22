// Students – one row per student
// Paste into Power Query: Home → Advanced Editor (after creating the DataFile parameter)
let
    Source = Excel.Workbook(File.Contents(DataFile), null, true),
    Students_Sheet = Source{[Item = "Students", Kind = "Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(Students_Sheet, [PromoteAllScalars = true]),
    // 12 messy spellings → 5 faculties
    #"Trimmed Text" = Table.TransformColumns(#"Promoted Headers", {{"Faculty", Text.Trim, type text}}),
    #"Capitalized Each Word" = Table.TransformColumns(#"Trimmed Text", {{"Faculty", Text.Proper, type text}}),
    #"Changed Type" = Table.TransformColumnTypes(
        #"Capitalized Each Word",
        {{"StudentID", type text}, {"StudentName", type text}, {"Batch", Int64.Type}, {"Hostel", type text}}
    )
in
    #"Changed Type"
