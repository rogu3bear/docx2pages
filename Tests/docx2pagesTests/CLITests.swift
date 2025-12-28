import XCTest
import Foundation

/// Integration tests for the docx2pages CLI.
/// These tests invoke the compiled binary and verify CLI behavior.
final class CLITests: XCTestCase {

    /// Path to the built binary
    var binaryPath: String {
        // During `swift test`, the binary is in .build/debug/
        let debugPath = URL(fileURLWithPath: #file)
            .deletingLastPathComponent()
            .deletingLastPathComponent()
            .deletingLastPathComponent()
            .appendingPathComponent(".build/debug/docx2pages")
        return debugPath.path
    }

    /// Run the CLI with arguments and return (stdout, stderr, exitCode)
    func runCLI(_ args: [String]) throws -> (stdout: String, stderr: String, exitCode: Int32) {
        let process = Process()
        process.executableURL = URL(fileURLWithPath: binaryPath)
        process.arguments = args

        let stdoutPipe = Pipe()
        let stderrPipe = Pipe()
        process.standardOutput = stdoutPipe
        process.standardError = stderrPipe

        try process.run()
        process.waitUntilExit()

        let stdoutData = stdoutPipe.fileHandleForReading.readDataToEndOfFile()
        let stderrData = stderrPipe.fileHandleForReading.readDataToEndOfFile()

        let stdout = String(data: stdoutData, encoding: .utf8) ?? ""
        let stderr = String(data: stderrData, encoding: .utf8) ?? ""

        return (stdout, stderr, process.terminationStatus)
    }

    // MARK: - Version Tests

    func testVersionFlag() throws {
        // Skip if binary doesn't exist (not built yet)
        guard FileManager.default.fileExists(atPath: binaryPath) else {
            throw XCTSkip("Binary not built yet. Run 'swift build' first.")
        }

        let result = try runCLI(["--version"])

        // Version should be in output (either stdout or combined)
        let combined = result.stdout + result.stderr
        XCTAssertTrue(combined.contains("1."), "Version output should contain version number")
    }

    // MARK: - Help Tests

    func testHelpFlag() throws {
        guard FileManager.default.fileExists(atPath: binaryPath) else {
            throw XCTSkip("Binary not built yet. Run 'swift build' first.")
        }

        let result = try runCLI(["--help"])

        let combined = result.stdout + result.stderr
        XCTAssertTrue(combined.contains("--input"), "Help should mention --input flag")
        XCTAssertTrue(combined.contains("--output"), "Help should mention --output flag")
        XCTAssertTrue(combined.contains("--template"), "Help should mention --template flag")
        XCTAssertTrue(combined.contains("--force-text"), "Help should mention --force-text flag")
        XCTAssertTrue(combined.contains("--timeout"), "Help should mention --timeout flag")
    }

    // MARK: - Validation Tests

    func testMissingRequiredArgs() throws {
        guard FileManager.default.fileExists(atPath: binaryPath) else {
            throw XCTSkip("Binary not built yet. Run 'swift build' first.")
        }

        let result = try runCLI([])

        // Should fail with non-zero exit code
        XCTAssertNotEqual(result.exitCode, 0, "Should fail without required args")

        // Should mention missing argument
        let combined = result.stdout + result.stderr
        XCTAssertTrue(
            combined.contains("--input") || combined.contains("Missing"),
            "Error should mention missing input argument"
        )
    }

    func testInvalidBatchSize() throws {
        guard FileManager.default.fileExists(atPath: binaryPath) else {
            throw XCTSkip("Binary not built yet. Run 'swift build' first.")
        }

        // Test batch-size below minimum
        let result = try runCLI([
            "--input", "/tmp/fake.docx",
            "--output", "/tmp/fake.pages",
            "--template", "/tmp/fake.pages",
            "--batch-size", "0"
        ])

        XCTAssertNotEqual(result.exitCode, 0, "Should fail with invalid batch-size")
        let combined = result.stdout + result.stderr
        XCTAssertTrue(
            combined.contains("batch-size") || combined.contains("between"),
            "Error should mention batch-size constraint"
        )
    }

    func testInvalidTimeout() throws {
        guard FileManager.default.fileExists(atPath: binaryPath) else {
            throw XCTSkip("Binary not built yet. Run 'swift build' first.")
        }

        // Test timeout above maximum
        let result = try runCLI([
            "--input", "/tmp/fake.docx",
            "--output", "/tmp/fake.pages",
            "--template", "/tmp/fake.pages",
            "--timeout", "9999"
        ])

        XCTAssertNotEqual(result.exitCode, 0, "Should fail with invalid timeout")
        let combined = result.stdout + result.stderr
        XCTAssertTrue(
            combined.contains("timeout") || combined.contains("between"),
            "Error should mention timeout constraint"
        )
    }

    func testNonexistentInputFile() throws {
        guard FileManager.default.fileExists(atPath: binaryPath) else {
            throw XCTSkip("Binary not built yet. Run 'swift build' first.")
        }

        let result = try runCLI([
            "--input", "/tmp/definitely-does-not-exist-12345.docx",
            "--output", "/tmp/out.pages",
            "--template", "/tmp/template.pages"
        ])

        XCTAssertNotEqual(result.exitCode, 0, "Should fail with nonexistent input")
        let combined = result.stdout + result.stderr
        XCTAssertTrue(
            combined.lowercased().contains("not found") || combined.lowercased().contains("does not exist"),
            "Error should mention file not found"
        )
    }

    func testWrongInputExtension() throws {
        guard FileManager.default.fileExists(atPath: binaryPath) else {
            throw XCTSkip("Binary not built yet. Run 'swift build' first.")
        }

        // Create temporary files - both input and template need to exist for extension check
        let tempDir = FileManager.default.temporaryDirectory
        let inputFile = tempDir.appendingPathComponent("test_\(UUID().uuidString).txt")
        let templateFile = tempDir.appendingPathComponent("test_\(UUID().uuidString).pages")

        try "test".write(to: inputFile, atomically: true, encoding: .utf8)
        try "test".write(to: templateFile, atomically: true, encoding: .utf8)
        defer {
            try? FileManager.default.removeItem(at: inputFile)
            try? FileManager.default.removeItem(at: templateFile)
        }

        let result = try runCLI([
            "--input", inputFile.path,
            "--output", "/tmp/out.pages",
            "--template", templateFile.path
        ])

        XCTAssertNotEqual(result.exitCode, 0, "Should fail with wrong extension")
        let combined = result.stdout + result.stderr
        XCTAssertTrue(
            combined.contains(".docx") || combined.contains("Input must be"),
            "Error should mention .docx requirement"
        )
    }
}
