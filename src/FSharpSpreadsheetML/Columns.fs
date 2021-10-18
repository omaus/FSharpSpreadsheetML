module Columns

open FSharpSpreadsheetML
open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Spreadsheet

/// Functions for working with Columns.
module Columns =

    /// Functions for working with a single Column.
    module Column =
        
        /// Creates an empty Column.
        let empty () = Column()

        /// Adds a given Column to a given Columns.
        let addColumn (columns : Columns) (column : Column) = columns.AppendChild column

        /// Sets a Column's BestFit property.
        let setBestFit (isBestFit : bool) (column : Column) = column.BestFit <- BooleanValue isBestFit; column

        /// Sets a Column's CustomWidth property.
        let setCustomWidth (isCustomWidth : bool) (column : Column) = column.CustomWidth <- BooleanValue istCustomWidth; column

        /// <summary>Sets a Column's width.</summary>
        /// <param name="width">The width of the Column to set to in (roughly) centimeters.</param>
        /// <param name="column">The column whose width to be set.</param>
        /// <remarks>The width value does not resemble a perfect change in centimeters. E.g., a width of 10. will result in the Column's width when opening the file in MS Excel to about 9.76 centimeters.</remarks>
        let setWidth (width : float) (column : Column) = column.Width <- DoubleValue width; column

        /// Sets a Column's index (1-based).
        let setIndex (index : uint) (column : Column) = 
            column.Min <- UInt32Value index
            column.Max <- UInt32Value index
            column

        /// Sets a Column's indices (1-based). The properties of this Column will then apply to all MS Excel columns with their indices being within this Column's index-span.
        let setIndices (leftIndex: uint) (rightIndex : uint) (column : Column) =
            column.Min <- UInt32Value leftIndex
            column.Max <- UInt32Value rightIndex
            column

        /// Sets a Column's Collapsed property.
        let setCollapsed (isCollapsed : bool) (column : Column) = column.Collapsed <- BooleanValue isCollapsed; column

        /// Sets a Column's Hidden property.
        let setHidden (isHidden : bool) (column : Column) = column.Hidden <- BooleanValue isHidden; column

        /// Sets a Column's OutlineLevel property.
        let setOutlineLevel (outlineLevel : byte) (column : Column) = column.OutlineLevel <- ByteValue outlineLevel

        /// Sets a Column's Phonetic property.
        let setPhonetic (isPhonetic : bool) (column : Column) = column.Phonetic <- BooleanValue isPhonetic

        /// Sets a Column's Style property.
        let setStyle (styleIndex : uint) (column : Column) = column.Style <- UInt32Value styleIndex

        /// Gets a Column's BestFit property.
        let getBestFit (column : Column) = column.BestFit.Value

        /// Gets a Column's CustomWidth property.
        let getCustomWidth (column : Column) = column.CustomWidth.Value

        /// <summary>Gets a Column's width.</summary>
        /// <param name="column">The column whose width to be get.</param>
        /// <remarks>The width value does not resemble a perfect change in centimeters. E.g., a width of 10. will result in the Column's width when opening the file in MS Excel to about 9.76 centimeters.</remarks>
        let getWidth (column : Column) = column.Width.Value

        /// Gets a Column's indices (1-based) (left and right index).
        let getIndices (column : Column) = column.Min.Value, column.Max.Value

        /// Gets a Column's Collapsed property.
        let getCollapsed (column : Column) = column.Collapsed.Value

        /// Gets a Column's Hidden property.
        let getHidden (column : Column) = column.Hidden.Value

        /// Gets a Column's OutlineLevel property.
        let getOutlineLevel (column : Column) = column.OutlineLevel.Value

        /// Gets a Column's Phonetic property.
        let getPhonetic (column : Column) = column.Phonetic.Value

        // Probably the index of a Style object referenced somewhere else...
        /// Gets a Column's Style property.
        let getStyle (column : Column) = column.Style.Value

        /// Creates a Column with the given index and width.
        let create index width =
            empty ()
            |> setIndex index
            |> setWidth width

        /// Creates a Column with the given width, spanning the given indices.
        let create2 leftIndex rightIndex width =
            empty ()
            |> setIndices leftIndex rightIndex
            |> setWidth width

        /// Creates a Column with the given index and width, and adds it to a given Columns.
        let init index width (columns : Columns) =
            create index width
            |> addColumn columns

        /// Creates a Column with the given width, spanning the given indices, and adds it to a given Columns.
        let init2 leftIndex rightIndex width (columns : Columns) =
            create2 leftIndex rightIndex width
            |> addColumn columns

    /// Creates an empty Columns.
    let empty () = Columns()

    /// Adds a given Column to a given Columns.
    let addColumn column columns = Column.addColumn columns column |> ignore; columns

    /// Adds a Column collection to a given Columns.
    let addColumns (columnSeq : seq<Column>) (columns : Columns) = columnSeq |> Seq.iter (Column.addColumn columns >> ignore); columns

    // Note: It seems that the Columns object must be placed BEFORE the SheetData object because otherwise the xlsx file will be corrupted.
    /// Adds a Columns to the Worksheet.
    let addColumnsToWorksheet (worksheet : Worksheet) (columns : Columns) = worksheet.InsertAt(columns, 0)

    let create column = empty () |> addColumn column