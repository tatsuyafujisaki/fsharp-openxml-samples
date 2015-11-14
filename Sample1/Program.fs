open System
open System.Linq
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Wordprocessing

[<EntryPoint>]
let main _ =
    use wd = WordprocessingDocument.Open(@"C:\foo.docx", false)
    wd.MainDocumentPart
        .Document
        .Body
        .Elements<Paragraph>()
        .FirstOrDefault()
        .FirstChild
        .ElementsAfter()
        |> Seq.filter (fun element -> element :? Run)
        |> Seq.iter (fun run -> printfn "%A" run.InnerText)

    Console.ReadKey() |> ignore
    0