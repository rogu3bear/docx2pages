#!/usr/bin/env osascript -l JavaScript
/**
 * HTML to Pages Exporter
 * Opens HTML in Pages and exports to .pages format.
 * Used as macOS 26+ workaround where JXA style APIs are broken.
 *
 * On macOS 26+, pages.open() returns null for direct file opens.
 * Workaround: Convert HTML to RTF (textutil), open RTF in Pages.
 *
 * Usage: osascript html_to_pages.js <input.html> <output.pages>
 */

function run(argv) {
    if (argv.length < 2) {
        return JSON.stringify({
            success: false,
            error: "Usage: html_to_pages.js <input.html> <output.pages>"
        });
    }

    const htmlPath = argv[0];
    const pagesPath = argv[1];

    // Generate temp RTF path
    const pid = $.NSProcessInfo.processInfo.processIdentifier;
    const rtfPath = `/tmp/docx2pages_temp_${pid}.rtf`;

    const app = Application.currentApplication();
    app.includeStandardAdditions = true;

    try {
        // Step 1: Convert HTML to RTF using textutil
        // This is needed because pages.open() is broken on macOS 26+
        // but RTF files open correctly via shell 'open' command
        try {
            app.doShellScript(`textutil -convert rtf "${htmlPath}" -output "${rtfPath}"`);
        } catch (e) {
            return JSON.stringify({
                success: false,
                error: "Failed to convert HTML to RTF: " + String(e)
            });
        }

        const pages = Application('Pages');
        pages.activate();

        // Get initial document count
        const initialCount = pages.documents.length;

        // Step 2: Open RTF in Pages using shell command
        // JXA pages.open() returns null on macOS 26+, but this works
        app.doShellScript(`open -a Pages "${rtfPath}"`);

        // Wait for document to open
        let attempts = 0;
        while (pages.documents.length <= initialCount && attempts < 50) {
            delay(0.1);
            attempts++;
        }

        if (pages.documents.length <= initialCount) {
            // Clean up temp file
            try { app.doShellScript(`rm -f "${rtfPath}"`); } catch(e) {}
            return JSON.stringify({
                success: false,
                error: "Timeout waiting for Pages to open RTF file"
            });
        }

        // Brief delay for full document load
        delay(0.5);

        // Step 3: Export as Pages format
        const doc = pages.documents[0];
        doc.export({
            to: Path(pagesPath),
            as: "Pages 09"
        });

        // Step 4: Close without saving (we already exported)
        doc.close({ saving: "no" });

        // Clean up temp RTF file
        try { app.doShellScript(`rm -f "${rtfPath}"`); } catch(e) {}

        return JSON.stringify({
            success: true,
            output: pagesPath
        });

    } catch (e) {
        // Clean up temp file on error
        try { app.doShellScript(`rm -f "${rtfPath}"`); } catch(err) {}
        return JSON.stringify({
            success: false,
            error: String(e)
        });
    }
}
