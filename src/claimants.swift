import Foundation
#if canImport(FoundationXML)
import FoundationXML
#endif
import AppKit

enum ClaimantError: LocalizedError {
    case message(String)
    var errorDescription: String? {
        switch self { case .message(let text): return text }
    }
}

struct ColumnMapping {
    let answerLetter: String
    let answerColumn: Int
    let headers: [String]
    let columns: [Int]

    static let standard = ColumnMapping(answerLetter: "E", answerColumn: 5,
                                        headers: ["Admin", "Admin P#", "Publisher", "Publisher P#"],
                                        columns: [18, 17, 14, 13])

    static func columnNumber(_ letters: String) -> Int? {
        let text = letters.trimmingCharacters(in: .whitespaces).uppercased()
        guard !text.isEmpty else { return nil }
        var number = 0
        for scalar in text.unicodeScalars {
            guard scalar.value >= 65 && scalar.value <= 90 else { return nil }
            number = number * 26 + Int(scalar.value - 64)
        }
        return number
    }

    // CLAIMANTS_COLUMNS names a settings file; otherwise claimant-columns.json beside the tool is used if present.
    static func load() throws -> ColumnMapping {
        var settingsURL: URL?
        if let path = ProcessInfo.processInfo.environment["CLAIMANTS_COLUMNS"], !path.isEmpty {
            settingsURL = URL(fileURLWithPath: path)
        } else if let beside = Bundle.main.executableURL?.deletingLastPathComponent().appendingPathComponent("claimant-columns.json"),
                  FileManager.default.fileExists(atPath: beside.path) {
            settingsURL = beside
        }
        guard let url = settingsURL else { return standard }
        guard let data = try? Data(contentsOf: url) else {
            throw ClaimantError.message("Could not read column settings: \(url.lastPathComponent).")
        }
        guard let object = (try? JSONSerialization.jsonObject(with: data)) as? [String: Any],
              let answerLetter = object["answerColumn"] as? String,
              let answerColumn = columnNumber(answerLetter),
              let fields = object["fields"] as? [[String: Any]], !fields.isEmpty else {
            throw ClaimantError.message("Column settings need an answerColumn letter and at least one field.")
        }
        var headers: [String] = []
        var columns: [Int] = []
        for field in fields {
            guard let header = field["header"] as? String,
                  let letters = field["column"] as? String,
                  let column = columnNumber(letters) else {
                throw ClaimantError.message("Each column setting needs a header and a column letter.")
            }
            headers.append(header)
            columns.append(column)
        }
        return ColumnMapping(answerLetter: answerLetter.uppercased(), answerColumn: answerColumn,
                             headers: headers, columns: columns)
    }
}

final class SharedStringsParser: NSObject, XMLParserDelegate {
    var strings: [String] = []
    private var inItem = false
    private var inText = false
    private var current = ""

    func parser(_ parser: XMLParser, didStartElement elementName: String,
                namespaceURI: String?, qualifiedName qName: String?,
                attributes attributeDict: [String: String] = [:]) {
        if elementName == "si" { inItem = true; current = "" }
        if inItem && elementName == "t" { inText = true }
    }

    func parser(_ parser: XMLParser, foundCharacters string: String) {
        if inItem && inText { current += string }
    }

    func parser(_ parser: XMLParser, didEndElement elementName: String,
                namespaceURI: String?, qualifiedName qName: String?) {
        if elementName == "t" { inText = false }
        if elementName == "si" { strings.append(current); inItem = false }
    }
}

final class WorkbookParser: NSObject, XMLParserDelegate {
    var firstRelationshipID: String?
    func parser(_ parser: XMLParser, didStartElement elementName: String,
                namespaceURI: String?, qualifiedName qName: String?,
                attributes attributeDict: [String: String] = [:]) {
        guard elementName == "sheet", firstRelationshipID == nil else { return }
        firstRelationshipID = attributeDict["r:id"] ?? attributeDict["id"]
    }
}

final class RelationshipsParser: NSObject, XMLParserDelegate {
    let wantedID: String
    var target: String?
    init(wantedID: String) { self.wantedID = wantedID }
    func parser(_ parser: XMLParser, didStartElement elementName: String,
                namespaceURI: String?, qualifiedName qName: String?,
                attributes attributeDict: [String: String] = [:]) {
        if elementName == "Relationship", attributeDict["Id"] == wantedID {
            target = attributeDict["Target"]
        }
    }
}

final class SheetParser: NSObject, XMLParserDelegate {
    let sharedStrings: [String]
    var cells: [Int: [Int: String]] = [:]
    private var cellReference = ""
    private var cellType = ""
    private var valueBuffer = ""
    private var inlineBuffer = ""
    private var collectingValue = false
    private var collectingInlineText = false
    private let wantedColumns: Set<Int>

    init(sharedStrings: [String], wantedColumns: Set<Int>) {
        self.sharedStrings = sharedStrings
        self.wantedColumns = wantedColumns
    }

    func parser(_ parser: XMLParser, didStartElement elementName: String,
                namespaceURI: String?, qualifiedName qName: String?,
                attributes attributeDict: [String: String] = [:]) {
        if elementName == "c" {
            cellReference = attributeDict["r"] ?? ""
            cellType = attributeDict["t"] ?? ""
            valueBuffer = ""
            inlineBuffer = ""
        } else if elementName == "v" {
            collectingValue = true
        } else if elementName == "t", cellType == "inlineStr" {
            collectingInlineText = true
        }
    }

    func parser(_ parser: XMLParser, foundCharacters string: String) {
        if collectingValue { valueBuffer += string }
        if collectingInlineText { inlineBuffer += string }
    }

    func parser(_ parser: XMLParser, didEndElement elementName: String,
                namespaceURI: String?, qualifiedName qName: String?) {
        if elementName == "v" { collectingValue = false }
        if elementName == "t" { collectingInlineText = false }
        if elementName == "c" {
            guard let (row, column) = Self.coordinates(cellReference), wantedColumns.contains(column) else { return }
            var value = cellType == "inlineStr" ? inlineBuffer : valueBuffer
            if cellType == "s", let index = Int(value), sharedStrings.indices.contains(index) {
                value = sharedStrings[index]
            } else if cellType == "b" {
                value = value == "1" ? "TRUE" : "FALSE"
            }
            cells[row, default: [:]][column] = value
        }
    }

    private static func coordinates(_ reference: String) -> (Int, Int)? {
        var column = 0
        var rowText = ""
        for scalar in reference.unicodeScalars {
            if scalar.value >= 65 && scalar.value <= 90 {
                column = column * 26 + Int(scalar.value - 64)
            } else if scalar.value >= 97 && scalar.value <= 122 {
                column = column * 26 + Int(scalar.value - 96)
            } else if scalar.value >= 48 && scalar.value <= 57 {
                rowText.unicodeScalars.append(scalar)
            }
        }
        guard column > 0, let row = Int(rowText) else { return nil }
        return (row, column)
    }
}

func parseXML(_ url: URL, delegate: XMLParserDelegate) throws {
    guard let parser = XMLParser(contentsOf: url) else {
        throw ClaimantError.message("Could not read \(url.lastPathComponent).")
    }
    parser.delegate = delegate
    guard parser.parse() else {
        throw parser.parserError ?? ClaimantError.message("Invalid workbook XML.")
    }
}

func resolveFirstSheet(in root: URL) throws -> URL {
    let xl = root.appendingPathComponent("xl")
    let workbookURL = xl.appendingPathComponent("workbook.xml")
    let workbookParser = WorkbookParser()
    try parseXML(workbookURL, delegate: workbookParser)

    guard let relationshipID = workbookParser.firstRelationshipID else {
        return xl.appendingPathComponent("worksheets/sheet1.xml")
    }

    let relationshipsURL = xl.appendingPathComponent("_rels/workbook.xml.rels")
    let relationshipsParser = RelationshipsParser(wantedID: relationshipID)
    try parseXML(relationshipsURL, delegate: relationshipsParser)
    guard var target = relationshipsParser.target else {
        return xl.appendingPathComponent("worksheets/sheet1.xml")
    }
    if target.hasPrefix("/") { target.removeFirst() }
    if target.hasPrefix("xl/") { return root.appendingPathComponent(target) }
    return xl.appendingPathComponent(target).standardizedFileURL
}

func csvField(_ value: String) -> String {
    if value.contains(",") || value.contains("\"") || value.contains("\n") || value.contains("\r") {
        return "\"" + value.replacingOccurrences(of: "\"", with: "\"\"") + "\""
    }
    return value
}

func processWorkbook(at inputURL: URL) throws -> String {
    guard inputURL.pathExtension.lowercased() == "xlsx" else {
        throw ClaimantError.message("Please drop an .xlsx report onto the app.")
    }
    let mapping = try ColumnMapping.load()

    let fileManager = FileManager.default
    let temporary = fileManager.temporaryDirectory.appendingPathComponent("Claimants-\(UUID().uuidString)", isDirectory: true)
    try fileManager.createDirectory(at: temporary, withIntermediateDirectories: true)
    defer { try? fileManager.removeItem(at: temporary) }

    let unzip = Process()
    unzip.executableURL = URL(fileURLWithPath: "/usr/bin/ditto")
    unzip.arguments = ["-x", "-k", inputURL.path, temporary.path]
    try unzip.run()
    unzip.waitUntilExit()
    guard unzip.terminationStatus == 0 else {
        throw ClaimantError.message("The workbook could not be opened.")
    }

    var sharedStrings: [String] = []
    let sharedURL = temporary.appendingPathComponent("xl/sharedStrings.xml")
    if fileManager.fileExists(atPath: sharedURL.path) {
        let sharedParser = SharedStringsParser()
        try parseXML(sharedURL, delegate: sharedParser)
        sharedStrings = sharedParser.strings
    }

    let sheetParser = SheetParser(sharedStrings: sharedStrings,
                                  wantedColumns: Set(mapping.columns + [mapping.answerColumn]))
    try parseXML(try resolveFirstSheet(in: temporary), delegate: sheetParser)

    let columns = mapping.columns
    let headers = mapping.headers

    var uniqueRows: [[String]] = []
    var seen = Set<[String]>()
    var hasNo = false
    var unknownRows: [Int] = []
    var matched = 0
    let lastRow = sheetParser.cells.keys.max() ?? 1
    if lastRow >= 2 {
        for row in 2...lastRow {
            let rowCells = sheetParser.cells[row] ?? [:]
            var answer = (rowCells[mapping.answerColumn] ?? "").trimmingCharacters(in: .whitespacesAndNewlines).uppercased()
            answer = answer.replacingOccurrences(of: "(", with: "").replacingOccurrences(of: ")", with: "")
            answer = answer.trimmingCharacters(in: .whitespacesAndNewlines)
            if answer.isEmpty { continue }
            if answer == "NO" || answer == "N" { hasNo = true; continue }
            guard answer == "YES" || answer == "Y" else { unknownRows.append(row); continue }
            matched += 1
            let claimant = columns.map { rowCells[$0] ?? "" }
            if seen.insert(claimant).inserted { uniqueRows.append(claimant) }
        }
    }

    guard matched > 0 else {
        throw ClaimantError.message("No YES, Yes, Y, or (Y) rows were found in column \(mapping.answerLetter).")
    }

    let classification = hasNo ? "True Partial" : (unknownRows.isEmpty ? "True Full" : "Needs review")
    let warning = unknownRows.isEmpty ? "" : " Unrecognized transfer answers in rows: \(unknownRows.map(String.init).joined(separator: ", "))."
    let allRows = [headers] + uniqueRows
    let tabSeparated = allRows.map { $0.joined(separator: "\t") }.joined(separator: "\n")

    let csv = allRows.map { $0.map(csvField).joined(separator: ",") }.joined(separator: "\r\n") + "\r\n"
    let baseName = inputURL.deletingPathExtension().lastPathComponent
    let outputURL = inputURL.deletingLastPathComponent().appendingPathComponent("\(baseName)_yes_claimants.csv")
    try csv.write(to: outputURL, atomically: true, encoding: .utf8)

    if ProcessInfo.processInfo.environment["CLAIMANTS_NO_CLIPBOARD"] != "1" {
        NSPasteboard.general.clearContents()
        NSPasteboard.general.setString(tabSeparated, forType: .string)
    }


    return "\(classification).\(warning) Found \(matched) matching rows and \(uniqueRows.count) unique claimant row\(uniqueRows.count == 1 ? "" : "s"). The table is copied to your clipboard and saved as:\n\(outputURL.lastPathComponent)"
}

do {
    guard CommandLine.arguments.count == 2 else {
        throw ClaimantError.message("Drop one .xlsx report onto the app.")
    }
    print(try processWorkbook(at: URL(fileURLWithPath: CommandLine.arguments[1])))
} catch {
    FileHandle.standardError.write(Data((error.localizedDescription + "\n").utf8))
    exit(1)
}
