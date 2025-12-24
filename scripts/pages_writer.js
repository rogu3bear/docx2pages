#!/usr/bin/env osascript -l JavaScript
/**
 * Pages Writer for docx2pages
 * Consumes a JSON block list and writes content into a Pages document.
 *
 * Usage: osascript -l JavaScript pages_writer.js --doc <output.pages> --json <blocks.json> [options]
 *
 * Options:
 *   --doc <path>           Path to Pages document to open and modify (required)
 *   --json <path>          Path to JSON blocks file (required)
 *   --strict               Fail if any style pollution detected
 *   --prefix-deep-headings Prefix deep headings with "HN:"
 *   --verbose              Enable verbose logging
 *   --table-style <name>   Table style name to apply
 *   --batch-size <n>       Flush paragraphs every N items (default: 50)
 */

ObjC.import('Foundation');

// Default batch size for chunked flushing
const DEFAULT_BATCH_SIZE = 50;

function run(argv) {
    // Parse arguments
    const args = parseArgs(argv);

    if (!args.doc || !args.json) {
        return JSON.stringify({
            success: false,
            error: "Usage: pages_writer.js --doc <output.pages> --json <blocks.json> [--strict] [--prefix-deep-headings] [--verbose] [--batch-size <n>]"
        });
    }

    const options = {
        strict: args.strict || false,
        prefixDeepHeadings: args['prefix-deep-headings'] || false,
        verbose: args.verbose || false,
        tableStyle: args['table-style'] || null,
        batchSize: parseInt(args['batch-size']) || DEFAULT_BATCH_SIZE
    };

    const log = (msg) => {
        if (options.verbose) {
            console.log(`[pages_writer] ${msg}`);
        }
    };

    // Read blocks JSON
    const blocksData = $.NSString.stringWithContentsOfFileEncodingError(args.json, $.NSUTF8StringEncoding, null);
    if (!blocksData) {
        return JSON.stringify({ success: false, error: `Cannot read blocks file: ${args.json}` });
    }

    let parsed;
    try {
        parsed = JSON.parse(ObjC.unwrap(blocksData));
    } catch (e) {
        return JSON.stringify({ success: false, error: `Invalid JSON in blocks file: ${e}` });
    }

    const blocks = parsed.blocks || [];

    // Track document reference for cleanup
    let doc = null;
    let baselineStyleNames = [];

    try {
        const pages = Application('Pages');
        pages.includeStandardAdditions = true;

        // Ensure Pages is running
        if (!pages.running()) {
            pages.activate();
            delay(1);
        }

        log(`Opening document: ${args.doc}`);

        // Open the document (this is already a COPY of the template, made by Swift)
        const docFile = Path(args.doc);
        doc = pages.open(docFile);
        delay(0.5);

        // Capture baseline styles IMMEDIATELY after open
        const baselineStyles = {};
        try {
            const styles = doc.paragraphStyles();
            for (let i = 0; i < styles.length; i++) {
                const styleName = styles[i].name();
                baselineStyles[styleName.toLowerCase()] = styleName;
                baselineStyleNames.push(styleName);
            }
        } catch (e) {
            log(`Warning: Could not enumerate styles: ${e}`);
        }

        log(`Baseline styles: ${baselineStyleNames.join(', ')}`);

        // Build style mapping from baseline styles
        const styleMap = buildStyleMap(baselineStyles, options);
        log(`Style mapping: ${JSON.stringify(styleMap)}`);

        // Detect list styles availability
        const listStyles = detectListStyles(baselineStyles);
        log(`List styles: bullet=${listStyles.bullet || 'none'}, numbered=${listStyles.numbered || 'none'}`);

        // PERFORMANCE: Cache style object references
        const styleCache = buildStyleCache(doc, baselineStyleNames, log);
        log(`Cached ${Object.keys(styleCache).length} style references`);

        log("Clearing body content...");

        // Clear body content
        try {
            const body = doc.bodyText;
            const textLength = body.characters.length;
            if (textLength > 0) {
                body.characters.slice(0, textLength).delete();
            }
        } catch (e) {
            log(`Note: Could not clear body via characters, trying alternative: ${e}`);
            try {
                doc.bodyText.set('');
            } catch (e2) {
                log(`Warning: Could not clear body: ${e2}`);
            }
        }

        delay(0.3);

        // Process blocks
        log(`Processing ${blocks.length} blocks (batch size: ${options.batchSize})...`);
        const results = {
            success: true,
            headingsWritten: 0,
            paragraphsWritten: 0,
            listsWritten: 0,
            tablesWritten: 0,
            tableFallbackCount: 0,
            warnings: [],
            stylesUsed: new Set(),
            listStyleUsed: false,
            paragraphErrors: 0
        };

        // Buffer for chunked paragraph writing
        const paragraphBuffer = [];

        for (let i = 0; i < blocks.length; i++) {
            const block = blocks[i];

            try {
                switch (block.type) {
                    case 'title':
                        bufferHeading(paragraphBuffer, block.text, 'title', styleMap, results, options, log);
                        results.headingsWritten++;
                        break;

                    case 'subtitle':
                        bufferHeading(paragraphBuffer, block.text, 'subtitle', styleMap, results, options, log);
                        results.headingsWritten++;
                        break;

                    case 'heading':
                        bufferHeading(paragraphBuffer, block.text, block.level, styleMap, results, options, log);
                        results.headingsWritten++;
                        break;

                    case 'paragraph':
                        bufferParagraph(paragraphBuffer, block.text, styleMap, results, log);
                        results.paragraphsWritten++;
                        break;

                    case 'list':
                        bufferList(paragraphBuffer, block, styleMap, listStyles, results, log);
                        results.listsWritten++;
                        break;

                    case 'table':
                        // Tables must be flushed and written immediately
                        flushParagraphBuffer(doc, paragraphBuffer, styleCache, results, options, log);
                        writeTable(doc, block, styleMap, options, results, log);
                        results.tablesWritten++;
                        break;

                    case 'break':
                        // Preserved break - buffer empty paragraph
                        bufferParagraph(paragraphBuffer, '', styleMap, results, log);
                        break;

                    default:
                        results.warnings.push(`Unknown block type: ${block.type}`);
                }

                // Chunked flushing: flush when buffer reaches batch size
                if (paragraphBuffer.length >= options.batchSize) {
                    log(`Flushing batch at block ${i + 1}...`);
                    flushParagraphBuffer(doc, paragraphBuffer, styleCache, results, options, log);
                }

            } catch (e) {
                results.warnings.push(`Error processing block ${i} (${block.type}): ${e}`);
                log(`Error: ${e}`);
            }
        }

        // Flush remaining paragraphs
        flushParagraphBuffer(doc, paragraphBuffer, styleCache, results, options, log);

        // Report paragraph styling errors if any
        if (results.paragraphErrors > 0) {
            const msg = `Failed to apply styles to ${results.paragraphErrors} paragraph(s)`;
            if (options.strict) {
                // In strict mode, paragraph errors are failures
                log("Closing document after strict failure (paragraph errors)...");
                doc.close({ saving: 'yes' });
                doc = null;

                return JSON.stringify({
                    success: false,
                    error: `Strict mode: ${msg}`,
                    paragraphErrors: results.paragraphErrors,
                    baselineStyles: baselineStyleNames,
                    finalStyles: [],
                    pollutingStyles: [],
                    unusedBaselineStyles: []
                });
            } else {
                results.warnings.push(msg);
            }
        }

        // Save document explicitly
        log("Saving document...");
        doc.save();
        delay(0.5);

        // Collect final style names to check for pollution
        const finalStyles = [];
        try {
            const styles = doc.paragraphStyles();
            for (let i = 0; i < styles.length; i++) {
                finalStyles.push(styles[i].name());
            }
        } catch (e) {
            log(`Warning: Could not enumerate final styles: ${e}`);
        }

        // Compute polluting styles (in final but not in baseline)
        const baselineSet = new Set(baselineStyleNames);
        const pollutingStyles = finalStyles.filter(s => !baselineSet.has(s));

        // Compute unused baseline styles (in baseline but not used)
        const usedSet = results.stylesUsed;
        const unusedBaselineStyles = baselineStyleNames.filter(s => !usedSet.has(s));

        if (pollutingStyles.length > 0) {
            const msg = `Style pollution detected: ${pollutingStyles.join(', ')}`;
            if (options.strict) {
                // In strict mode, still save and close properly before returning error
                log("Closing document after strict failure...");
                doc.close({ saving: 'yes' });
                doc = null;

                return JSON.stringify({
                    success: false,
                    error: msg,
                    baselineStyles: baselineStyleNames,
                    finalStyles: finalStyles,
                    pollutingStyles: pollutingStyles,
                    unusedBaselineStyles: unusedBaselineStyles
                });
            } else {
                results.warnings.push(msg);
            }
        }

        // Close document properly
        log("Closing document...");
        doc.close({ saving: 'yes' });
        doc = null;

        log("Done!");

        return JSON.stringify({
            success: true,
            headingsWritten: results.headingsWritten,
            paragraphsWritten: results.paragraphsWritten,
            listsWritten: results.listsWritten,
            tablesWritten: results.tablesWritten,
            tableFallbackCount: results.tableFallbackCount,
            warnings: results.warnings,
            baselineStyles: baselineStyleNames,
            finalStyles: finalStyles,
            stylesUsed: Array.from(results.stylesUsed),
            pollutingStyles: pollutingStyles,
            unusedBaselineStyles: unusedBaselineStyles,
            listStyleUsed: results.listStyleUsed
        });

    } catch (e) {
        // Cleanup: try to close document on error
        if (doc !== null) {
            try {
                doc.close({ saving: 'no' });
            } catch (closeErr) {
                // Ignore close errors during cleanup
            }
        }
        return JSON.stringify({
            success: false,
            error: String(e),
            stack: e.stack,
            baselineStyles: baselineStyleNames,
            finalStyles: [],
            pollutingStyles: [],
            unusedBaselineStyles: []
        });
    }
}

function parseArgs(argv) {
    const args = {};
    let i = 0;
    while (i < argv.length) {
        const arg = argv[i];
        if (arg.startsWith('--')) {
            const key = arg.slice(2);
            // Check if next arg is a value or another flag
            if (i + 1 < argv.length && !argv[i + 1].startsWith('--')) {
                args[key] = argv[i + 1];
                i += 2;
            } else {
                args[key] = true;
                i += 1;
            }
        } else {
            i += 1;
        }
    }
    return args;
}

function buildStyleMap(baselineStyles, options) {
    const map = {
        body: null,
        title: null,
        subtitle: null,
        headings: {}
    };

    // Body style - preference order
    for (const pref of ['body', 'body text', 'normal']) {
        if (baselineStyles[pref]) {
            map.body = baselineStyles[pref];
            break;
        }
    }

    // Title style
    if (baselineStyles['title']) {
        map.title = baselineStyles['title'];
    }

    // Subtitle style
    if (baselineStyles['subtitle']) {
        map.subtitle = baselineStyles['subtitle'];
    }

    // Heading styles - find max available level
    let maxHeadingLevel = 0;

    // Check for "Heading" (level 1)
    if (baselineStyles['heading']) {
        map.headings[1] = baselineStyles['heading'];
        maxHeadingLevel = 1;
    }

    // Check for "Heading N" patterns
    for (let level = 1; level <= 9; level++) {
        const key = `heading ${level}`;
        if (baselineStyles[key]) {
            map.headings[level] = baselineStyles[key];
            maxHeadingLevel = Math.max(maxHeadingLevel, level);
        }
    }

    // Also check "Heading" alone for level 1 if not already set
    if (!map.headings[1] && baselineStyles['heading']) {
        map.headings[1] = baselineStyles['heading'];
        maxHeadingLevel = Math.max(maxHeadingLevel, 1);
    }

    map.maxHeadingLevel = maxHeadingLevel;

    // Fallbacks
    if (!map.title) {
        map.title = map.headings[1] || map.body;
    }
    if (!map.subtitle) {
        map.subtitle = map.body;
    }

    return map;
}

function detectListStyles(baselineStyles) {
    const listStyles = {
        bullet: null,
        numbered: null
    };

    // Bullet list style - preference order
    const bulletPrefs = ['bullet', 'bulleted', 'bulleted list', 'bullets'];
    for (const pref of bulletPrefs) {
        if (baselineStyles[pref]) {
            listStyles.bullet = baselineStyles[pref];
            break;
        }
    }

    // Numbered list style - preference order
    const numberedPrefs = ['numbered', 'numbered list', 'numbers'];
    for (const pref of numberedPrefs) {
        if (baselineStyles[pref]) {
            listStyles.numbered = baselineStyles[pref];
            break;
        }
    }

    return listStyles;
}

/**
 * Build a cache of style object references to avoid repeated lookups.
 * This significantly improves performance for large documents.
 */
function buildStyleCache(doc, styleNames, log) {
    const cache = {};

    for (const styleName of styleNames) {
        try {
            const styles = doc.paragraphStyles.whose({ name: styleName });
            if (styles.length > 0) {
                cache[styleName] = styles[0];
            }
        } catch (e) {
            log(`Warning: Could not cache style '${styleName}': ${e}`);
        }
    }

    return cache;
}

/**
 * Convert column index to Excel-style letters.
 * 0 -> A, 25 -> Z, 26 -> AA, 27 -> AB, ..., 51 -> AZ, 52 -> BA, etc.
 */
function colIndexToLetters(c) {
    let result = '';
    let n = c;
    while (n >= 0) {
        result = String.fromCharCode(65 + (n % 26)) + result;
        n = Math.floor(n / 26) - 1;
    }
    return result;
}

// PERFORMANCE: Buffer-based paragraph writing with chunked flushing
// Paragraphs are collected and written in batches to balance memory and IPC overhead.

function bufferHeading(buffer, text, levelOrType, styleMap, results, options, log) {
    if (!text || text.trim() === '') return;

    let styleName;
    let displayText = text;

    if (levelOrType === 'title') {
        styleName = styleMap.title;
    } else if (levelOrType === 'subtitle') {
        styleName = styleMap.subtitle;
    } else {
        const level = parseInt(levelOrType);
        const effectiveLevel = Math.min(level, styleMap.maxHeadingLevel || 1);

        if (options.prefixDeepHeadings && level > styleMap.maxHeadingLevel) {
            displayText = `H${level}: ${text}`;
        }

        styleName = styleMap.headings[effectiveLevel] || styleMap.headings[1] || styleMap.body;
    }

    buffer.push({ text: displayText, style: styleName });
}

function bufferParagraph(buffer, text, styleMap, results, log) {
    // Empty paragraphs are preserved for spacing
    buffer.push({ text: text || '', style: styleMap.body });
}

function bufferList(buffer, block, styleMap, listStyles, results, log) {
    const items = block.items;
    const ordered = block.ordered;

    // Determine which style to use
    const listStyle = ordered ? listStyles.numbered : listStyles.bullet;
    const usingNativeStyle = listStyle !== null;

    if (usingNativeStyle) {
        results.listStyleUsed = true;
    }

    for (let i = 0; i < items.length; i++) {
        const item = items[i];
        const text = typeof item === 'string' ? item : item.text;
        const level = typeof item === 'object' ? (item.level || 0) : 0;

        let lineText;

        if (usingNativeStyle) {
            // Use native list style - indent with tabs for nesting
            const indent = '\t'.repeat(level);
            lineText = indent + text;
        } else {
            // Fallback: manual prefix
            const indent = '    '.repeat(level);
            let prefix;
            if (ordered) {
                prefix = `${i + 1}. `;
            } else {
                prefix = '• ';
            }
            lineText = indent + prefix + text;
        }

        buffer.push({ text: lineText, style: usingNativeStyle ? listStyle : styleMap.body });
    }

    if (!usingNativeStyle) {
        results.warnings.push('List written as formatted text (no list paragraph style in template)');
    }
}

/**
 * Flush buffered paragraphs to the document using O(n) insertion.
 *
 * PERFORMANCE FIX: Instead of reading the entire body and rewriting it (O(n²)),
 * we use insertionPoints[-1] to insert at the end without reading existing content.
 * If that fails, we fall back to appending, but track paragraphsBefore to only
 * apply styles to new paragraphs.
 */
function flushParagraphBuffer(doc, buffer, styleCache, results, options, log) {
    if (buffer.length === 0) return;

    log(`Flushing ${buffer.length} paragraphs...`);

    // Build full text with newlines
    const fullText = buffer.map(p => p.text).join('\n') + '\n';

    const body = doc.bodyText;

    // Track paragraph count BEFORE insertion for style application
    let paragraphsBefore = 0;
    try {
        paragraphsBefore = body.paragraphs.length;
    } catch (e) {
        // If we can't count, assume 0
    }

    // O(n) insertion: insert at end without reading entire body
    // Uses insertionPoints to avoid the read-all/write-all pattern
    try {
        // Try insertion point approach (O(batch_size), not O(total_document))
        const insertionPoint = body.insertionPoints[-1];
        insertionPoint.contents = fullText;
    } catch (e) {
        // Fallback: direct set for empty document, or append for non-empty
        log(`InsertionPoint failed, using fallback: ${e}`);
        try {
            if (paragraphsBefore === 0) {
                body.set(fullText);
            } else {
                // Last resort: read and append (O(n²) but rare)
                const currentText = body();
                body.set(currentText + fullText);
            }
        } catch (e2) {
            log(`Warning: Could not write text: ${e2}`);
            results.paragraphErrors += buffer.length;
            buffer.length = 0;
            return;
        }
    }

    // Small delay for Pages to process
    delay(0.1);

    // Apply styles to NEW paragraphs only (using tracked startIndex)
    const paragraphs = body.paragraphs;
    const totalParagraphs = paragraphs.length;
    const startIndex = paragraphsBefore;

    for (let i = 0; i < buffer.length; i++) {
        const entry = buffer[i];
        if (entry.style) {
            try {
                const paraIndex = startIndex + i;
                if (paraIndex >= 0 && paraIndex < totalParagraphs) {
                    // Use cached style reference
                    const cachedStyle = styleCache[entry.style];
                    if (cachedStyle) {
                        paragraphs[paraIndex].paragraphStyle = cachedStyle;
                        results.stylesUsed.add(entry.style);
                    }
                }
            } catch (e) {
                // Per-paragraph error recovery: continue with other paragraphs
                results.paragraphErrors = (results.paragraphErrors || 0) + 1;
            }
        }
    }

    // Clear buffer
    buffer.length = 0;
}

function writeTable(doc, block, styleMap, options, results, log) {
    const rows = block.rows;
    if (!rows || rows.length === 0) return;

    const numRows = rows.length;
    const numCols = Math.max(...rows.map(r => r.length));

    log(`Creating table: ${numRows} rows x ${numCols} cols`);

    try {
        const table = doc.tables.push(doc.Table({
            rowCount: numRows,
            columnCount: numCols
        }));

        delay(0.2);

        // Fill cells using Excel-style column addressing (supports >26 columns)
        // Process in row chunks for better progress on very large tables
        const ROW_CHUNK_SIZE = 50;

        for (let rStart = 0; rStart < numRows; rStart += ROW_CHUNK_SIZE) {
            const rEnd = Math.min(rStart + ROW_CHUNK_SIZE, numRows);

            for (let r = rStart; r < rEnd; r++) {
                for (let c = 0; c < rows[r].length; c++) {
                    try {
                        const colLetters = colIndexToLetters(c);
                        const cellAddr = `${colLetters}${r + 1}`;
                        const cell = table.cells[cellAddr];
                        cell.value = rows[r][c] || '';
                    } catch (e) {
                        log(`Warning: Could not set cell [${r},${c}]: ${e}`);
                    }
                }
            }

            // Small delay between chunks for very large tables
            if (rEnd < numRows && numRows > ROW_CHUNK_SIZE) {
                delay(0.05);
            }
        }

        // Add newline after table by appending to body
        const body = doc.bodyText;
        const currentText = body();
        body.set(currentText + '\n');

    } catch (e) {
        log(`Error creating table, falling back to text: ${e}`);
        results.tableFallbackCount++;
        results.warnings.push(`Table ${results.tablesWritten + 1} fell back to text: ${e}`);

        // Fallback: write as text
        const body = doc.bodyText;
        let fallbackText = '';
        for (const row of rows) {
            fallbackText += '| ' + row.join(' | ') + ' |\n';
        }
        fallbackText += '\n';

        const currentText = body();
        body.set(currentText + fallbackText);
    }
}
