open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Wordprocessing

let createTableBorders size =
    let size' = UInt32Value (uint32 size)
    let tbs = new TableBorders()
    tbs.AppendChild(new TopBorder (Val = new EnumValue<BorderValues>(BorderValues.Single), Size = size')) |> ignore
    tbs.AppendChild(new BottomBorder (Val = new EnumValue<BorderValues>(BorderValues.Single), Size = size')) |> ignore
    tbs.AppendChild(new LeftBorder (Val = new EnumValue<BorderValues>(BorderValues.Single), Size = size')) |> ignore
    tbs.AppendChild(new RightBorder (Val = new EnumValue<BorderValues>(BorderValues.Single), Size = size')) |> ignore
    tbs.AppendChild(new InsideHorizontalBorder (Val = new EnumValue<BorderValues>(BorderValues.Single), Size = size')) |> ignore
    tbs.AppendChild(new InsideVerticalBorder (Val = new EnumValue<BorderValues>(BorderValues.Single), Size = size')) |> ignore
    tbs

type E = OpenXmlElement

let createTable (data : string[,]) borderSize =
    let tps = new TableProperties()
    let tbs = createTableBorders borderSize
    tps.AppendChild(tbs) |> ignore

    let table = new Table(tps :> OpenXmlElement)
    for i = 0 to data.GetUpperBound(0) do
        let tr = new TableRow()
        for j = 0 to data.GetUpperBound(1) do
            tr.Append(new TableCell(new Paragraph(new Run(new Text(data.[i, j]) :> E) :> E) :> E,
                        new TableCellProperties(new TableCellWidth (Width = StringValue "100", Type = EnumValue TableWidthUnitValues.Pct)
                        :> E) :> E) :> E)
        table.Append(tr :> E)
    table

[<EntryPoint>]
let main _ = 
    let filePath = @"C:\x\result.docx"
    let text = "Test of creating a table."
    let data = array2D [|[| "a"; "b" |]; [| "c"; "d" |]|]

    use wd = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document)
    wd.AddMainDocumentPart().Document <- new Document(new Body(new Paragraph(new Run(new Text(text) :> E) :> E) :> E) :> E) 

    let table = createTable data 1

    let doc = wd.MainDocumentPart.Document
    doc.Body.AppendChild(table) |> ignore
    doc.Save()
    0
