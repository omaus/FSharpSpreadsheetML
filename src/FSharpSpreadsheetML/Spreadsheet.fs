namespace FSharpSpreadsheetML

open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Spreadsheet



/// Functions for working the spreadsheet document.
module Spreadsheet = 

    /// Opens the SpreadsheetDocument located at the given path and initialized a FileStream.
    let fromFile (path : string) isEditable = SpreadsheetDocument.Open(path,isEditable)

    /// Opens the SpreadsheetDocument from the given FileStream.
    let fromStream (stream : System.IO.Stream) isEditable = SpreadsheetDocument.Open(stream,isEditable)

    /// Initializes a new empty SpreadsheetDocument at the given path.
    let initEmpty (path : string) = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook)

    /// Gets the WorkbookPart of the SpreadsheetDocument.
    let getWorkbookPart (spreadsheet : SpreadsheetDocument) = spreadsheet.WorkbookPart

    /// Gets the Workbook of the SpreadsheetDocument's WorkbookPart.
    let getWorkbook (spreadsheet : SpreadsheetDocument) = getWorkbookPart spreadsheet |> Workbook.get

    // Only if none there
    /// Initialized a new workbookPart in the spreadsheetDocument but only if there is none.
    let initWorkbookPart (spreadsheet : SpreadsheetDocument) = spreadsheet.AddWorkbookPart()

    /// Saves changes made to the spreadsheet.
    let saveChanges (spreadsheet : SpreadsheetDocument) = 
        spreadsheet.Save() 
        spreadsheet

    /// Closes the FileStream to the spreadsheet.
    let close (spreadsheet : SpreadsheetDocument) = spreadsheet.Close()

    /// Saves changes made to the spreadsheet to the given path.
    let saveAs path (spreadsheet : SpreadsheetDocument) = 
        spreadsheet.SaveAs(path) :?> SpreadsheetDocument
        |> close
        spreadsheet

    /// Initializes a new spreadsheet with an empty sheet at the given path.
    let init sheetName (path : string) = 
        let doc = initEmpty path
        let workbookPart = initWorkbookPart doc

        WorkbookPart.appendSheet sheetName (SheetData.empty ()) workbookPart |> ignore
        doc

    /// Initializes a new spreadsheet with an empty sheet and a sharedStringTable at the given path.
    let initWithSst sheetName (path : string) = 
        let doc = init sheetName path
        let workbookPart = getWorkbookPart doc

        let sharedStringTablePart = WorkbookPart.getOrInitSharedStringTablePart workbookPart
        SharedStringTable.init sharedStringTablePart |> ignore

        doc

    // Gets the SharedStringTable from the SharedStringTablePart of a SpreadsheetDocument's WorkbookPart.
    let getSharedStringTable (spreadsheetDocument : SpreadsheetDocument) =
        spreadsheetDocument.WorkbookPart.SharedStringTablePart.SharedStringTable

    // Gets the SharedStringTable from the SharedStringTablePart of a SpreadsheetDocument's WorkbookPart if it exists, else returns None.
    let tryGetSharedStringTable (spreadsheetDocument : SpreadsheetDocument) =
        try spreadsheetDocument.WorkbookPart.SharedStringTablePart.SharedStringTable |> Some
        with | _ -> None

    // Gets the SharedStringTablePart of a SpreadsheetDocument's WorkbookPart. If it does not exist, creates a new one.
    let getOrInitSharedStringTablePart (spreadsheetDocument : SpreadsheetDocument) =
        let workbookPart = spreadsheetDocument.WorkbookPart    
        let sstp = workbookPart.GetPartsOfType<SharedStringTablePart>()
        match sstp |> Seq.tryHead with
        | Some sst -> sst
        | None -> workbookPart.AddNewPart<SharedStringTablePart>()

    /// Returns the WorksheetPart associated to the sheet with the given name if it exists. Else returns None.
    let tryGetWorksheetPartBySheetName (name : string) (spreadsheetDocument : SpreadsheetDocument) =
        Sheet.tryItemByName name spreadsheetDocument
        |> Option.map (fun sheet -> 
            let sheetId = Sheet.getID sheet
            getWorkbookPart spreadsheetDocument
            |> Worksheet.WorksheetPart.getByID sheetId
        )

    /// Returns the WorksheetPart associated to the sheet with the given name.
    let getWorksheetPartBySheetName (name : string) (spreadsheetDocument : SpreadsheetDocument) =
        tryGetWorksheetPartBySheetName name spreadsheetDocument |> Option.get

    /// Returns the Worksheet associated to the sheet with the given name if it exists. Else returns None.
    let tryGetWorksheetBySheetName (name : string) (spreadsheetDocument : SpreadsheetDocument) =
        match tryGetWorksheetPartBySheetName name spreadsheetDocument with
        | Some wsp  -> Some (Worksheet.get wsp)
        | None      -> None

    /// Returns the Worksheet associated to the sheet with the given name.
    let getWorksheetBySheetName (name : string) (spreadsheetDocument : SpreadsheetDocument) =
        tryGetWorksheetBySheetName name spreadsheetDocument |> Option.get
    
    /// Returns the WorksheetPart for the given 0-based sheetIndex of the given SpreadsheetDocument if it exists. Else returns None.
    let tryGetWorksheetPartBySheetIndex (sheetIndex : uint) (spreadsheetDocument : SpreadsheetDocument) =
        Sheet.tryItem sheetIndex spreadsheetDocument
        |> Option.map (fun sheet -> 
            spreadsheetDocument.WorkbookPart
            |> Worksheet.WorksheetPart.getByID sheet.Id.Value 
        )

    /// Returns the WorksheetPart for the given 0-based sheetIndex of the given SpreadsheetDocument. 
    let getWorksheetPartBySheetIndex (sheetIndex : uint) (spreadsheetDocument : SpreadsheetDocument) =
        tryGetWorksheetPartBySheetIndex sheetIndex spreadsheetDocument |> Option.get

    /// Returns the Worksheet for the given 0-based sheetIndex of the given SpreadsheetDocument if it exists. Else returns None.
    let tryGetWorksheetBySheetIndex (sheetIndex : uint) (spreadsheetDocument : SpreadsheetDocument) =
        match tryGetWorksheetPartBySheetIndex sheetIndex spreadsheetDocument with
        | Some wsp  -> Some (Worksheet.get wsp)
        | None      -> None

    /// Returns the Worksheet for the given 0-based sheetIndex of the given SpreadsheetDocument.
    let getWorksheetBySheetIndex (sheetIndex : uint) (spreadsheetDocument : SpreadsheetDocument) =
        tryGetWorksheetBySheetIndex sheetIndex spreadsheetDocument |> Option.get
        
    /// Returns the SheetData for the given Sheet's name of the given SpreadsheetDocument if it exists. Else returns None.
    let tryGetSheetBySheetName (name : string) (spreadsheetDocument : SpreadsheetDocument) =
        tryGetWorksheetPartBySheetName name spreadsheetDocument
        |> Option.map (Worksheet.get >> Worksheet.getSheetData)

    /// Returns the SheetData for the given Sheet's name of the given SpreadsheetDocument.
    let getSheetBySheetName (name : string) (spreadsheetDocument : SpreadsheetDocument) =
        tryGetSheetBySheetName name spreadsheetDocument |> Option.get

    /// Returns the SheetData for the given 0-based sheetIndex of the given SpreadsheetDocument if it exists. Else returns None.
    let tryGetSheetBySheetIndex (sheetIndex : uint) (spreadsheetDocument : SpreadsheetDocument) =
        tryGetWorksheetPartBySheetIndex sheetIndex spreadsheetDocument
        |> Option.map (Worksheet.get >> Worksheet.getSheetData)

    /// Returns the SheetData for the given 0-based sheetIndex of the given SpreadsheetDocument.
    let getSheetBySheetIndex (sheetIndex : uint) (spreadsheetDocument : SpreadsheetDocument) =
        tryGetSheetBySheetIndex sheetIndex spreadsheetDocument |> Option.get
        
    let tryGetSheetNameBySheetIndex (sheetIndex : uint) (spreadsheetDocument : SpreadsheetDocument) =
        match Sheet.tryItem sheetIndex spreadsheetDocument with
        | Some sheet    -> Some (Sheet.getName sheet)
        | None          -> None
    
    /// Returns the name of the Sheet with the given index of the given SpreadsheetDocument.
    let getSheetNameBySheetIndex (sheetIndex : uint) (spreadsheetDocument : SpreadsheetDocument) = 
        (tryGetSheetNameBySheetIndex sheetIndex spreadsheetDocument).Value

    /// Returns a sequence of rows containing the cells for the given 0-based sheetIndex of the given SpreadsheetDocument. 
    let getRowsBySheetIndex (sheetIndex : uint) (spreadsheetDocument : SpreadsheetDocument) =

        match (Sheet.tryItem sheetIndex spreadsheetDocument) with
        | Some (sheet) ->
            let workbookPart = getWorkbookPart spreadsheetDocument
            let sheedId = Sheet.getID sheet
            let worksheetPart = Worksheet.WorksheetPart.getByID sheedId workbookPart     
            let stringTablePart = getOrInitSharedStringTablePart spreadsheetDocument
            seq {
            use reader = OpenXmlReader.Create(worksheetPart)
      
            while reader.Read() do
                if (reader.ElementType = typeof<Row>) then 
                    let row = reader.LoadCurrentElement() :?> Row
                    row.Elements()
                    |> Seq.iter (fun item -> 
                        let cell = item :?> Cell
                        Cell.includeSharedStringValue stringTablePart.SharedStringTable cell |> ignore
                        )
                    yield row 
            }
        | None -> Seq.empty

    /// Returns a 1D-sequence of cells for the given sheetIndex of the given spreadsheetDocument. 
    let getCellsBySheetIndex (sheetIndex : uint) (spreadsheetDocument : SpreadsheetDocument) =

        match (Sheet.tryItem sheetIndex spreadsheetDocument) with
        | Some (sheet) ->
            let workbookPart = spreadsheetDocument.WorkbookPart
            let worksheetPart = Worksheet.WorksheetPart.getByID sheet.Id.Value workbookPart
            let stringTablePart = getOrInitSharedStringTablePart spreadsheetDocument
            seq {
            use reader = OpenXmlReader.Create(worksheetPart)
        
            while reader.Read() do
                if (reader.ElementType = typeof<Cell>) then 
                    let cell    = reader.LoadCurrentElement() :?> Cell 
                    let cellRef = if cell.CellReference.HasValue then cell.CellReference.Value else ""
                    yield Cell.includeSharedStringValue stringTablePart.SharedStringTable cell
            }
        | None -> seq {()}

    //----------------------------------------------------------------------------------------------------------------------
    //                                      High level functions                                                            
    //----------------------------------------------------------------------------------------------------------------------

    //Rows

    let mapRowOfSheet (sheetId) (rowId) (rowF: Row -> Row) : SpreadsheetDocument = 
        //get workbook part
        //get sheet data by sheetId
        //get row at rowId
        //apply rowF to row and update 
        //return updated doc
        raise (System.NotImplementedException())

    let mapRowsOfSheet (sheetId) (rowF: Row -> Row) : SpreadsheetDocument = raise (System.NotImplementedException())

    let appendRowValuesToSheet (sheetId) (rowValues: seq<'T>) : SpreadsheetDocument = raise (System.NotImplementedException())

    let insertRowValuesIntoSheetAt (sheetId) (rowId) (rowValues: seq<'T>) : SpreadsheetDocument = raise (System.NotImplementedException())

    let insertValueIntoSheetAt (sheetId) (rowId) (colId) (value: 'T) : SpreadsheetDocument = raise (System.NotImplementedException())

    let setValueInSheetAt (sheetId) (rowId) (colId) (value: 'T) : SpreadsheetDocument = raise (System.NotImplementedException())

    let deleteRowFromSheet (sheetId) (rowId) : SpreadsheetDocument = raise (System.NotImplementedException())

    //...






