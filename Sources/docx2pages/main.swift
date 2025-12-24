import Foundation
import ArgumentParser

@main
struct Docx2Pages: ParsableCommand {
    static let configuration = CommandConfiguration(
        commandName: "docx2pages",
        abstract: "Convert DOCX files to Pages documents using a template's native styles.",
        discussion: """
            This tool parses a DOCX file to extract document structure (headings, paragraphs,
            lists, tables) and creates a new Pages document from a template, applying the
            template's paragraph styles to preserve visual consistency.

            The output document uses only the template's styles - no Word style pollution.
            """,
        version: "1.2.0"
    )

    @Option(name: .shortAndLong, help: "Input DOCX file path")
    var input: String

    @Option(name: .shortAndLong, help: "Output .pages file path")
    var output: String

    @Option(name: .shortAndLong, help: "Pages template file path")
    var template: String

    @Flag(name: .long, help: "Prefix deep headings (beyond template max) with 'HN:'")
    var prefixDeepHeadings: Bool = false

    @Option(name: .long, help: "Table style name to apply from template")
    var tableStyle: String?

    @Flag(name: .long, help: "Fail on any style pollution or fallback behavior")
    var strict: Bool = false

    @Flag(name: .long, help: "Convert page/section breaks to blank paragraphs")
    var preserveBreaks: Bool = false

    @Flag(name: .long, help: "Overwrite output file if it exists")
    var overwrite: Bool = false

    @Flag(name: .shortAndLong, help: "Enable verbose logging")
    var verbose: Bool = false

    @Flag(name: .long, help: "Disable concurrency lock (advanced)")
    var noLock: Bool = false

    @Flag(name: .long, help: "Fail immediately if lock is held (don't wait)")
    var noWait: Bool = false

    @Option(name: .long, help: "Write JSON summary to file (use '-' for stdout)")
    var jsonSummary: String?

    @Option(name: .long, help: "Paragraph batch size for Pages writer (default: 50)")
    var batchSize: Int = 50

    @Option(name: .long, help: "Directory containing scripts (parse_docx.py, pages_writer.js)")
    var scriptsDir: String?

    // Lock file descriptor for concurrency safety
    private static var lockFd: Int32 = -1

    func run() throws {
        let startTime = Date()

        // Determine output destinations based on --json-summary
        let useStdoutForJson = jsonSummary == "-"
        let logOutput: (String) -> Void = useStdoutForJson
            ? { msg in FileHandle.standardError.write("\(msg)\n".data(using: .utf8)!) }
            : { msg in print(msg) }

        func log(_ message: String) {
            logOutput(message)
        }

        func verboseLog(_ message: String) {
            if verbose {
                logOutput("[verbose] \(message)")
            }
        }

        log("docx2pages v1.2.0")
        log("================")

        // Acquire lock unless --no-lock
        if !noLock {
            try acquireLock(blocking: !noWait)
            verboseLog("Acquired conversion lock")
        }

        defer {
            if !noLock {
                releaseLock()
            }
        }

        // Validate inputs
        let inputURL = URL(fileURLWithPath: input)
        let outputURL = URL(fileURLWithPath: output)
        let templateURL = URL(fileURLWithPath: template)

        guard FileManager.default.fileExists(atPath: inputURL.path) else {
            throw ValidationError("Input file not found: \(input)")
        }

        guard FileManager.default.fileExists(atPath: templateURL.path) else {
            throw ValidationError("Template file not found: \(template)")
        }

        guard input.lowercased().hasSuffix(".docx") else {
            throw ValidationError("Input must be a .docx file")
        }

        guard output.lowercased().hasSuffix(".pages") else {
            throw ValidationError("Output must be a .pages file")
        }

        guard template.lowercased().hasSuffix(".pages") else {
            throw ValidationError("Template must be a .pages file")
        }

        // Ensure output parent directory exists
        let outputParent = outputURL.deletingLastPathComponent()
        if !FileManager.default.fileExists(atPath: outputParent.path) {
            verboseLog("Creating output directory: \(outputParent.path)")
            do {
                try FileManager.default.createDirectory(at: outputParent, withIntermediateDirectories: true)
            } catch {
                throw ValidationError("Cannot create output directory '\(outputParent.path)': \(error.localizedDescription)")
            }
        }

        // Check output exists
        if FileManager.default.fileExists(atPath: outputURL.path) {
            if overwrite {
                verboseLog("Output exists, will overwrite: \(output)")
            } else {
                throw ValidationError("Output file already exists: \(output). Use --overwrite to replace.")
            }
        }

        // Check for Pages app
        guard let pagesPath = findPagesApp() else {
            throw ValidationError("Pages.app not found. Please install Pages from the App Store.")
        }
        verboseLog("Found Pages at: \(pagesPath)")

        log("Input:    \(input)")
        log("Output:   \(output)")
        log("Template: \(template)")
        if strict {
            log("Mode:     strict (fail on pollution/fallback)")
        }
        log("")

        // Step 1: Parse DOCX
        log("Step 1: Parsing DOCX...")
        let parseResult = try parseDocx(inputURL: inputURL, verboseLog: verboseLog)

        // In strict mode, fail on severe parser warnings
        if strict {
            let severeWarnings = parseResult.stats.warnings.filter { warning in
                // These indicate structural problems that make conversion unreliable
                warning.contains("Invalid ZIP") ||
                warning.contains("invalid DOCX") ||
                warning.contains("No word/document.xml") ||
                warning.contains("Malformed") ||
                warning.contains("No body element")
            }
            if !severeWarnings.isEmpty {
                throw ValidationError("Strict mode: Parser encountered severe issues: \(severeWarnings.joined(separator: "; "))")
            }

            // Also fail if document appears empty but file is not
            if parseResult.blocks.isEmpty {
                let fileSize = try FileManager.default.attributesOfItem(atPath: inputURL.path)[.size] as? Int ?? 0
                if fileSize > 1000 {  // Minimal DOCX is ~1KB
                    throw ValidationError("Strict mode: Parser returned zero blocks from a non-trivial file (\(fileSize) bytes)")
                }
            }
        }

        // Print parsing stats
        printParseStats(parseResult.stats, log: log)

        // Step 2: Copy template to TEMP output (atomic write pattern)
        log("")
        log("Step 2: Preparing output from template...")
        let tempOutputURL = try prepareTempOutput(from: templateURL, finalOutput: outputURL)
        verboseLog("Temp output: \(tempOutputURL.path)")

        // Step 3: Write to Pages (opens the temp copy)
        log("")
        log("Step 3: Writing to Pages document...")

        let writeResult: WriteResult

        do {
            writeResult = try writeToPagesDocument(
                blocks: parseResult,
                outputURL: tempOutputURL,
                verboseLog: verboseLog
            )
        } catch {
            // Clean up temp file on failure
            try? FileManager.default.removeItem(at: tempOutputURL)
            throw error
        }

        // Step 4: Validate results
        var strictFailures: [String] = []

        // Check style pollution
        if !writeResult.pollutingStyles.isEmpty {
            let msg = "Style pollution: \(writeResult.pollutingStyles.joined(separator: ", "))"
            if strict {
                strictFailures.append(msg)
            }
        }

        // Check table fallback
        if writeResult.tableFallbackCount > 0 {
            let msg = "Table fallback: \(writeResult.tableFallbackCount) table(s) rendered as text"
            if strict {
                strictFailures.append(msg)
            }
        }

        // Check list fallback (only if lists exist)
        let hasLists = parseResult.stats.lists.bulleted + parseResult.stats.lists.numbered > 0
        if hasLists && !writeResult.listStyleUsed {
            let msg = "List fallback: lists rendered as formatted text (no list styles in template)"
            if strict {
                strictFailures.append(msg)
            }
        }

        // Print write results
        printWriteResults(writeResult, strict: strict, log: log)

        // Step 5: Atomic move to final location
        let elapsed = Date().timeIntervalSince(startTime)
        var finalSuccess = true
        var finalError: String? = nil

        if strictFailures.isEmpty {
            log("")
            log("Step 4: Finalizing output...")
            try safeMove(from: tempOutputURL, to: outputURL, overwrite: overwrite)
            verboseLog("Safely moved to: \(output)")
        } else {
            // Clean up temp on strict failure
            try? FileManager.default.removeItem(at: tempOutputURL)

            log("")
            log("✗ Strict mode failures:")
            for failure in strictFailures {
                log("  • \(failure)")
            }
            finalSuccess = false
            finalError = "Strict mode: \(strictFailures.count) failure(s) detected"
        }

        log("")
        log(String(format: "Completed in %.2f seconds", elapsed))
        if finalSuccess {
            log("Output: \(output)")
        }

        // Write JSON summary if requested
        if let summaryPath = jsonSummary {
            let summary = JSONSummary(
                toolVersion: "1.2.0",
                input: input,
                output: output,
                template: template,
                strict: strict,
                parseStats: parseResult.stats,
                writeResult: writeResult,
                elapsedSeconds: elapsed,
                success: finalSuccess,
                error: finalError
            )

            let encoder = JSONEncoder()
            encoder.outputFormatting = [.prettyPrinted, .sortedKeys]
            let jsonData = try encoder.encode(summary)

            if summaryPath == "-" {
                // Write to stdout
                print(String(data: jsonData, encoding: .utf8)!)
            } else {
                // Write to file
                try jsonData.write(to: URL(fileURLWithPath: summaryPath))
                log("JSON summary written to: \(summaryPath)")
            }
        }

        if !finalSuccess {
            throw ValidationError(finalError!)
        }
    }

    private func findPagesApp() -> String? {
        let possiblePaths = [
            "/Applications/Pages.app",
            "/System/Applications/Pages.app",
            NSHomeDirectory() + "/Applications/Pages.app"
        ]

        for path in possiblePaths {
            if FileManager.default.fileExists(atPath: path) {
                return path
            }
        }

        return nil
    }

    private func acquireLock(blocking: Bool) throws {
        let lockPath = "/tmp/docx2pages.lock"
        let fd = open(lockPath, O_CREAT | O_RDWR, 0o644)
        guard fd >= 0 else {
            throw ValidationError("Could not create lock file: \(lockPath)")
        }

        // Try to acquire exclusive lock
        let lockFlags: Int32 = blocking ? LOCK_EX : (LOCK_EX | LOCK_NB)
        let result = flock(fd, lockFlags)

        if result != 0 {
            close(fd)
            if !blocking && errno == EWOULDBLOCK {
                throw ValidationError("Another docx2pages process is running. Use without --no-wait to wait, or use --no-lock to skip locking.")
            }
            throw ValidationError("Could not acquire lock. Another docx2pages process may be running.")
        }

        Docx2Pages.lockFd = fd
    }

    private func releaseLock() {
        if Docx2Pages.lockFd >= 0 {
            flock(Docx2Pages.lockFd, LOCK_UN)
            close(Docx2Pages.lockFd)
            Docx2Pages.lockFd = -1
        }
    }

    private func prepareTempOutput(from template: URL, finalOutput: URL) throws -> URL {
        let fm = FileManager.default

        // Create temp path next to final output
        let tempName = ".\(finalOutput.lastPathComponent).tmp"
        let tempURL = finalOutput.deletingLastPathComponent().appendingPathComponent(tempName)

        // Remove existing temp if present
        if fm.fileExists(atPath: tempURL.path) {
            try fm.removeItem(at: tempURL)
        }

        // Copy template to temp location
        try fm.copyItem(at: template, to: tempURL)

        return tempURL
    }

    private func safeMove(from source: URL, to destination: URL, overwrite: Bool) throws {
        let fm = FileManager.default

        if fm.fileExists(atPath: destination.path) && overwrite {
            // Safe overwrite: backup existing, move new, remove backup
            // This avoids the gap where destination is deleted but move hasn't happened
            let backupName = ".\(destination.lastPathComponent).backup"
            let backupURL = destination.deletingLastPathComponent().appendingPathComponent(backupName)

            // Remove any stale backup
            if fm.fileExists(atPath: backupURL.path) {
                try fm.removeItem(at: backupURL)
            }

            // Move existing to backup
            try fm.moveItem(at: destination, to: backupURL)

            do {
                // Move temp to destination
                try fm.moveItem(at: source, to: destination)
                // Success - remove backup
                try? fm.removeItem(at: backupURL)
            } catch {
                // Restore from backup on failure
                try? fm.moveItem(at: backupURL, to: destination)
                throw error
            }
        } else {
            // No existing file or not overwriting
            try fm.moveItem(at: source, to: destination)
        }
    }

    private func parseDocx(inputURL: URL, verboseLog: (String) -> Void) throws -> ParseResult {
        // Find the Python parser script
        let scriptPath = findScript("parse_docx.py")

        guard let scriptPath = scriptPath else {
            throw ValidationError("Could not find parse_docx.py script")
        }

        verboseLog("Using parser: \(scriptPath)")

        // Create temp file for output
        let tempDir = FileManager.default.temporaryDirectory
        let blocksFile = tempDir.appendingPathComponent("docx2pages_blocks_\(ProcessInfo.processInfo.processIdentifier).json")

        defer {
            try? FileManager.default.removeItem(at: blocksFile)
        }

        // Run Python parser
        let process = Process()
        process.executableURL = URL(fileURLWithPath: "/usr/bin/python3")
        process.arguments = [scriptPath, inputURL.path, "-o", blocksFile.path]

        if verbose {
            process.arguments?.append("-v")
        }

        if preserveBreaks {
            process.arguments?.append("--preserve-breaks")
        }

        let errorPipe = Pipe()
        process.standardError = errorPipe

        try process.run()
        process.waitUntilExit()

        if process.terminationStatus != 0 {
            let errorData = errorPipe.fileHandleForReading.readDataToEndOfFile()
            let errorMessage = String(data: errorData, encoding: .utf8) ?? "Unknown error"
            throw ValidationError("DOCX parsing failed: \(errorMessage)")
        }

        // Read the JSON output
        let jsonData = try Data(contentsOf: blocksFile)
        let result = try JSONDecoder().decode(ParseResult.self, from: jsonData)

        return result
    }

    private func writeToPagesDocument(blocks: ParseResult, outputURL: URL, verboseLog: (String) -> Void) throws -> WriteResult {
        // Find the JXA script
        let scriptPath = findScript("pages_writer.js")

        guard let scriptPath = scriptPath else {
            throw ValidationError("Could not find pages_writer.js script")
        }

        verboseLog("Using writer: \(scriptPath)")

        // Create temp file for blocks JSON
        let tempDir = FileManager.default.temporaryDirectory
        let pid = ProcessInfo.processInfo.processIdentifier
        let blocksFile = tempDir.appendingPathComponent("docx2pages_blocks_\(pid).json")

        defer {
            try? FileManager.default.removeItem(at: blocksFile)
        }

        // Write blocks JSON
        let blocksData = try JSONEncoder().encode(blocks)
        try blocksData.write(to: blocksFile)

        // Build arguments for pages_writer.js
        var args = [
            "-l", "JavaScript",
            scriptPath,
            "--doc", outputURL.path,
            "--json", blocksFile.path
        ]

        if strict {
            args.append("--strict")
        }

        if prefixDeepHeadings {
            args.append("--prefix-deep-headings")
        }

        if verbose {
            args.append("--verbose")
        }

        if let tableStyle = tableStyle {
            args.append("--table-style")
            args.append(tableStyle)
        }

        if batchSize != 50 {
            args.append("--batch-size")
            args.append(String(batchSize))
        }

        // Run JXA script via osascript
        let process = Process()
        process.executableURL = URL(fileURLWithPath: "/usr/bin/osascript")
        process.arguments = args

        let outputPipe = Pipe()
        let errorPipe = Pipe()
        process.standardOutput = outputPipe
        process.standardError = errorPipe

        verboseLog("Running Pages automation...")

        try process.run()
        process.waitUntilExit()

        let outputData = outputPipe.fileHandleForReading.readDataToEndOfFile()
        let errorData = errorPipe.fileHandleForReading.readDataToEndOfFile()

        if verbose {
            if let errorOutput = String(data: errorData, encoding: .utf8), !errorOutput.isEmpty {
                for line in errorOutput.components(separatedBy: "\n") {
                    if !line.isEmpty {
                        verboseLog("[pages] \(line)")
                    }
                }
            }
        }

        guard let outputString = String(data: outputData, encoding: .utf8)?.trimmingCharacters(in: .whitespacesAndNewlines),
              !outputString.isEmpty else {
            throw ValidationError("Pages automation returned no output")
        }

        // Parse result JSON
        guard let resultData = outputString.data(using: .utf8) else {
            throw ValidationError("Could not parse Pages result")
        }

        do {
            let result = try JSONDecoder().decode(WriteResult.self, from: resultData)
            if !result.success {
                throw ValidationError("Pages automation failed: \(result.error ?? "Unknown error")")
            }
            return result
        } catch {
            verboseLog("Raw output: \(outputString)")
            throw ValidationError("Failed to parse Pages result: \(error)")
        }
    }

    private func findScript(_ name: String) -> String? {
        var searchPaths: [String] = []

        // Priority 1: Explicit --scripts-dir if provided
        if let explicitDir = scriptsDir {
            let explicitPath = (explicitDir as NSString).appendingPathComponent(name)
            searchPaths.append(explicitPath)
        }

        // Priority 2: Relative to executable (for dist package: <exe>/../scripts/)
        if let executablePath = Bundle.main.executablePath {
            let execDir = (executablePath as NSString).deletingLastPathComponent
            let distPath = (execDir as NSString).appendingPathComponent("../scripts/\(name)")
            searchPaths.append(distPath)
        }

        // Priority 3: Development paths
        searchPaths.append(contentsOf: [
            // Development: relative to executable bundle
            Bundle.main.bundlePath + "/../scripts/\(name)",
            // Development: relative to working directory
            FileManager.default.currentDirectoryPath + "/scripts/\(name)",
            // Built product with resources
            Bundle.main.resourcePath.map { $0 + "/scripts/\(name)" },
            // Homebrew style installation
            "/usr/local/share/docx2pages/scripts/\(name)",
            // User installation
            NSHomeDirectory() + "/.docx2pages/scripts/\(name)"
        ].compactMap { $0 })

        for path in searchPaths {
            // Normalize path to resolve ../ components
            let normalizedPath = (path as NSString).standardizingPath
            if FileManager.default.fileExists(atPath: normalizedPath) {
                return normalizedPath
            }
        }

        return nil
    }

    private func printParseStats(_ stats: ParseStats, log: (String) -> Void) {
        log("")
        log("Document Structure:")
        log("-------------------")

        // Headings
        var totalHeadings = 0
        if let title = stats.headings["title"], title > 0 {
            log("  Title:      \(title)")
            totalHeadings += title
        }
        if let subtitle = stats.headings["subtitle"], subtitle > 0 {
            log("  Subtitle:   \(subtitle)")
            totalHeadings += subtitle
        }
        for level in 1...9 {
            let key = "level_\(level)"
            if let count = stats.headings[key], count > 0 {
                log("  Heading \(level):  \(count)")
                totalHeadings += count
            }
        }
        if totalHeadings > 0 {
            log("  Total:      \(totalHeadings) headings")
        }

        // Paragraphs
        if stats.paragraphs > 0 {
            log("  Paragraphs: \(stats.paragraphs)")
        }

        // Lists
        let totalLists = stats.lists.bulleted + stats.lists.numbered
        if totalLists > 0 {
            log("  Lists:      \(totalLists) (\(stats.lists.bulleted) bulleted, \(stats.lists.numbered) numbered)")
        }

        // Tables
        if stats.tables.count > 0 {
            log("  Tables:     \(stats.tables.count) (max \(stats.tables.maxRows)×\(stats.tables.maxCols))")
        }

        // Breaks dropped
        if stats.droppedBreaks > 0 {
            log("  Breaks:     \(stats.droppedBreaks) dropped")
        }

        // Warnings
        if !stats.warnings.isEmpty {
            log("")
            log("Warnings:")
            for warning in stats.warnings {
                log("  ⚠ \(warning)")
            }
        }
    }

    private func printWriteResults(_ result: WriteResult, strict: Bool, log: (String) -> Void) {
        log("")
        log("Pages Output:")
        log("-------------")
        log("  Headings:   \(result.headingsWritten)")
        log("  Paragraphs: \(result.paragraphsWritten)")
        log("  Lists:      \(result.listsWritten)")
        log("  Tables:     \(result.tablesWritten)")

        if result.listStyleUsed {
            log("  List mode:  native styles")
        } else if result.listsWritten > 0 {
            log("  List mode:  text fallback")
        }

        if result.tableFallbackCount > 0 {
            log("  Table mode: \(result.tableFallbackCount) fell back to text")
        }

        if !result.stylesUsed.isEmpty {
            log("")
            log("Styles Applied:")
            for style in result.stylesUsed.sorted() {
                log("  • \(style)")
            }
        }

        // Report unused baseline styles (informational)
        if !result.unusedBaselineStyles.isEmpty {
            log("")
            log("Unused Template Styles:")
            for style in result.unusedBaselineStyles.sorted().prefix(5) {
                log("  • \(style)")
            }
            if result.unusedBaselineStyles.count > 5 {
                log("  ... and \(result.unusedBaselineStyles.count - 5) more")
            }
        }

        // Check for style pollution
        if !result.pollutingStyles.isEmpty {
            log("")
            if strict {
                log("✗ Style Pollution (strict mode will fail):")
            } else {
                log("⚠ Style Pollution Detected:")
            }
            for style in result.pollutingStyles.sorted() {
                log("  • \(style)")
            }
        } else if !result.baselineStyles.isEmpty {
            log("")
            log("✓ No style pollution detected")
        }

        if !result.warnings.isEmpty {
            log("")
            log("Warnings:")
            for warning in result.warnings {
                log("  ⚠ \(warning)")
            }
        }
    }
}

// MARK: - Models

struct ParseResult: Codable {
    let blocks: [Block]
    let stats: ParseStats
}

struct Block: Codable {
    let type: String
    let text: String?
    let level: Int?
    let ordered: Bool?
    let items: [ListItem]?
    let rows: [[String]]?
}

struct ListItem: Codable {
    let text: String
    let level: Int?
}

struct ParseStats: Codable {
    let headings: [String: Int]
    let paragraphs: Int
    let lists: ListStats
    let tables: TableStats
    let warnings: [String]
    let droppedBreaks: Int

    enum CodingKeys: String, CodingKey {
        case headings
        case paragraphs
        case lists
        case tables
        case warnings
        case droppedBreaks = "dropped_breaks"
    }

    init(from decoder: Decoder) throws {
        let container = try decoder.container(keyedBy: CodingKeys.self)
        headings = try container.decode([String: Int].self, forKey: .headings)
        paragraphs = try container.decode(Int.self, forKey: .paragraphs)
        lists = try container.decode(ListStats.self, forKey: .lists)
        tables = try container.decode(TableStats.self, forKey: .tables)
        warnings = try container.decode([String].self, forKey: .warnings)
        droppedBreaks = try container.decodeIfPresent(Int.self, forKey: .droppedBreaks) ?? 0
    }
}

struct ListStats: Codable {
    let bulleted: Int
    let numbered: Int
}

struct TableStats: Codable {
    let count: Int
    let maxRows: Int
    let maxCols: Int

    enum CodingKeys: String, CodingKey {
        case count
        case maxRows = "max_rows"
        case maxCols = "max_cols"
    }
}

struct WriteResult: Codable {
    let success: Bool
    let error: String?
    let headingsWritten: Int
    let paragraphsWritten: Int
    let listsWritten: Int
    let tablesWritten: Int
    let tableFallbackCount: Int
    let warnings: [String]
    let baselineStyles: [String]
    let finalStyles: [String]
    let stylesUsed: [String]
    let pollutingStyles: [String]
    let unusedBaselineStyles: [String]
    let listStyleUsed: Bool

    init(from decoder: Decoder) throws {
        let container = try decoder.container(keyedBy: CodingKeys.self)
        success = try container.decode(Bool.self, forKey: .success)
        error = try container.decodeIfPresent(String.self, forKey: .error)
        headingsWritten = try container.decodeIfPresent(Int.self, forKey: .headingsWritten) ?? 0
        paragraphsWritten = try container.decodeIfPresent(Int.self, forKey: .paragraphsWritten) ?? 0
        listsWritten = try container.decodeIfPresent(Int.self, forKey: .listsWritten) ?? 0
        tablesWritten = try container.decodeIfPresent(Int.self, forKey: .tablesWritten) ?? 0
        tableFallbackCount = try container.decodeIfPresent(Int.self, forKey: .tableFallbackCount) ?? 0
        warnings = try container.decodeIfPresent([String].self, forKey: .warnings) ?? []
        baselineStyles = try container.decodeIfPresent([String].self, forKey: .baselineStyles) ?? []
        finalStyles = try container.decodeIfPresent([String].self, forKey: .finalStyles) ?? []
        stylesUsed = try container.decodeIfPresent([String].self, forKey: .stylesUsed) ?? []
        pollutingStyles = try container.decodeIfPresent([String].self, forKey: .pollutingStyles) ?? []
        unusedBaselineStyles = try container.decodeIfPresent([String].self, forKey: .unusedBaselineStyles) ?? []
        listStyleUsed = try container.decodeIfPresent(Bool.self, forKey: .listStyleUsed) ?? false
    }

    enum CodingKeys: String, CodingKey {
        case success
        case error
        case headingsWritten
        case paragraphsWritten
        case listsWritten
        case tablesWritten
        case tableFallbackCount
        case warnings
        case baselineStyles
        case finalStyles
        case stylesUsed
        case pollutingStyles
        case unusedBaselineStyles
        case listStyleUsed
    }
}

struct JSONSummary: Codable {
    let toolVersion: String
    let input: String
    let output: String
    let template: String
    let strict: Bool
    let parseStats: ParseStats
    let writeResult: WriteResultSummary
    let elapsedSeconds: Double
    let success: Bool
    let error: String?
}

struct WriteResultSummary: Codable {
    let headingsWritten: Int
    let paragraphsWritten: Int
    let listsWritten: Int
    let tablesWritten: Int
    let tableFallbackCount: Int
    let warnings: [String]
    let stylesUsed: [String]
    let pollutingStyles: [String]
    let unusedBaselineStyles: [String]
    let listStyleUsed: Bool

    init(from result: WriteResult) {
        self.headingsWritten = result.headingsWritten
        self.paragraphsWritten = result.paragraphsWritten
        self.listsWritten = result.listsWritten
        self.tablesWritten = result.tablesWritten
        self.tableFallbackCount = result.tableFallbackCount
        self.warnings = result.warnings
        self.stylesUsed = result.stylesUsed
        self.pollutingStyles = result.pollutingStyles
        self.unusedBaselineStyles = result.unusedBaselineStyles
        self.listStyleUsed = result.listStyleUsed
    }
}

extension JSONSummary {
    init(toolVersion: String, input: String, output: String, template: String,
         strict: Bool, parseStats: ParseStats, writeResult: WriteResult,
         elapsedSeconds: Double, success: Bool, error: String?) {
        self.toolVersion = toolVersion
        self.input = input
        self.output = output
        self.template = template
        self.strict = strict
        self.parseStats = parseStats
        self.writeResult = WriteResultSummary(from: writeResult)
        self.elapsedSeconds = elapsedSeconds
        self.success = success
        self.error = error
    }
}

struct ValidationError: Error, CustomStringConvertible {
    let message: String

    init(_ message: String) {
        self.message = message
    }

    var description: String { message }
}
