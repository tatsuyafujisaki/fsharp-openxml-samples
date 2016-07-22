open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Wordprocessing

let private createTableBorders size =
    let size' = UInt32Value (uint32 size)
    let tbs = TableBorders()
    tbs.AppendChild(TopBorder (Val = EnumValue<BorderValues>(BorderValues.Single), Size = size')) |> ignore
    tbs.AppendChild(BottomBorder (Val = EnumValue<BorderValues>(BorderValues.Single), Size = size')) |> ignore
    tbs.AppendChild(LeftBorder (Val = EnumValue<BorderValues>(BorderValues.Single), Size = size')) |> ignore
    tbs.AppendChild(RightBorder (Val = EnumValue<BorderValues>(BorderValues.Single), Size = size')) |> ignore
    tbs.AppendChild(InsideHorizontalBorder (Val = EnumValue<BorderValues>(BorderValues.Single), Size = size')) |> ignore
    tbs.AppendChild(InsideVerticalBorder (Val = EnumValue<BorderValues>(BorderValues.Single), Size = size')) |> ignore
    tbs

type private E = OpenXmlElement

let private createTable (data : string[,]) borderSize =
    let tps = TableProperties()
    let tbs = createTableBorders borderSize
    tps.AppendChild(tbs) |> ignore

    let table = Table(tps :> E)
    for i = 0 to data.GetUpperBound(0) do
        let tr = TableRow()
        for j = 0 to data.GetUpperBound(1) do
            tr.Append(TableCell(Paragraph(Run(Text(data.[i, j]) :> E) :> E) :> E,
                        TableCellProperties(TableCellWidth (Width = StringValue "100",
                                                            Type = EnumValue TableWidthUnitValues.Pct)
                        :> E) :> E) :> E)
        table.Append(tr :> E)
    table

[<EntryPoint>]
let private main _ = 
    let filePath = @"C:\result.docx"
    let text = "Test of creating a table."
    let data = array2D [|[| "a"; "b" |]; [| "c"; "d" |]|]

    use wd = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document)
    wd.AddMainDocumentPart().Document <- Document(Body(Paragraph(Run(Text(text) :> E) :> E) :> E) :> E) 

    let table = createTable data 1

    let doc = wd.MainDocumentPart.Document
    doc.Body.AppendChild(table) |> ignore
    doc.Save()
    0
