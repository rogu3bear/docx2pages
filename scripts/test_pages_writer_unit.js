#!/usr/bin/env node
/**
 * Unit tests for pure functions in pages_writer.js
 *
 * Tests utility functions that don't depend on JXA/Pages APIs.
 * Run with: node scripts/test_pages_writer_unit.js
 */

// =============================================================================
// COPY OF PURE FUNCTIONS FROM pages_writer.js
// These are copied to enable testing without JXA dependencies
// =============================================================================

function colIndexToLetters(c) {
    let result = '';
    let n = c;
    while (n >= 0) {
        result = String.fromCharCode(65 + (n % 26)) + result;
        n = Math.floor(n / 26) - 1;
    }
    return result;
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
    for (const pref of ['bullet', 'bulleted', 'bulleted list', 'bullets']) {
        if (baselineStyles[pref]) {
            listStyles.bullet = baselineStyles[pref];
            break;
        }
    }

    // Numbered list style - preference order
    for (const pref of ['numbered', 'numbered list', 'numbers']) {
        if (baselineStyles[pref]) {
            listStyles.numbered = baselineStyles[pref];
            break;
        }
    }

    return listStyles;
}

// =============================================================================
// TEST FRAMEWORK (minimal)
// =============================================================================

let passed = 0;
let failed = 0;

function assert(condition, message) {
    if (condition) {
        passed++;
        console.log(`  ✓ ${message}`);
    } else {
        failed++;
        console.log(`  ✗ ${message}`);
    }
}

function assertEqual(actual, expected, message) {
    if (actual === expected) {
        passed++;
        console.log(`  ✓ ${message}`);
    } else {
        failed++;
        console.log(`  ✗ ${message}`);
        console.log(`    Expected: ${expected}`);
        console.log(`    Actual: ${actual}`);
    }
}

function assertDeepEqual(actual, expected, message) {
    if (JSON.stringify(actual) === JSON.stringify(expected)) {
        passed++;
        console.log(`  ✓ ${message}`);
    } else {
        failed++;
        console.log(`  ✗ ${message}`);
        console.log(`    Expected: ${JSON.stringify(expected)}`);
        console.log(`    Actual: ${JSON.stringify(actual)}`);
    }
}

// =============================================================================
// TESTS
// =============================================================================

console.log('\ncolIndexToLetters tests:');

assertEqual(colIndexToLetters(0), 'A', 'Column 0 should be A');
assertEqual(colIndexToLetters(1), 'B', 'Column 1 should be B');
assertEqual(colIndexToLetters(25), 'Z', 'Column 25 should be Z');
assertEqual(colIndexToLetters(26), 'AA', 'Column 26 should be AA');
assertEqual(colIndexToLetters(27), 'AB', 'Column 27 should be AB');
assertEqual(colIndexToLetters(51), 'AZ', 'Column 51 should be AZ');
assertEqual(colIndexToLetters(52), 'BA', 'Column 52 should be BA');
assertEqual(colIndexToLetters(701), 'ZZ', 'Column 701 should be ZZ');
assertEqual(colIndexToLetters(702), 'AAA', 'Column 702 should be AAA');

console.log('\nbuildStyleMap tests:');

// Test with full style set
const fullStyles = {
    'body': 'Body',
    'title': 'Title',
    'subtitle': 'Subtitle',
    'heading': 'Heading',
    'heading 1': 'Heading 1',
    'heading 2': 'Heading 2',
    'heading 3': 'Heading 3'
};
const fullMap = buildStyleMap(fullStyles, {});
assertEqual(fullMap.body, 'Body', 'Body style should be found');
assertEqual(fullMap.title, 'Title', 'Title style should be found');
assertEqual(fullMap.subtitle, 'Subtitle', 'Subtitle style should be found');
assertEqual(fullMap.headings[1], 'Heading 1', 'Heading 1 should be found');
assertEqual(fullMap.headings[2], 'Heading 2', 'Heading 2 should be found');
assertEqual(fullMap.headings[3], 'Heading 3', 'Heading 3 should be found');
assertEqual(fullMap.maxHeadingLevel, 3, 'Max heading level should be 3');

// Test body fallback to "body text"
const bodyTextStyles = {
    'body text': 'Body Text'
};
const bodyTextMap = buildStyleMap(bodyTextStyles, {});
assertEqual(bodyTextMap.body, 'Body Text', 'Body Text fallback should work');

// Test body fallback to "normal"
const normalStyles = {
    'normal': 'Normal'
};
const normalMap = buildStyleMap(normalStyles, {});
assertEqual(normalMap.body, 'Normal', 'Normal fallback should work');

// Test title fallback to heading 1
const noTitleStyles = {
    'body': 'Body',
    'heading 1': 'Heading 1'
};
const noTitleMap = buildStyleMap(noTitleStyles, {});
assertEqual(noTitleMap.title, 'Heading 1', 'Title should fallback to Heading 1');

// Test empty styles
const emptyMap = buildStyleMap({}, {});
assertEqual(emptyMap.body, null, 'Body should be null with empty styles');
assertEqual(emptyMap.maxHeadingLevel, 0, 'Max heading level should be 0');

// Test "Heading" alone (without number)
const headingAloneStyles = {
    'body': 'Body',
    'heading': 'Heading'
};
const headingAloneMap = buildStyleMap(headingAloneStyles, {});
assertEqual(headingAloneMap.headings[1], 'Heading', 'Heading alone should map to level 1');
assertEqual(headingAloneMap.maxHeadingLevel, 1, 'Max level should be 1 with Heading alone');

console.log('\ndetectListStyles tests:');

// Test bullet style detection
const bulletStyles = {
    'bullet': 'Bullet'
};
const bulletResult = detectListStyles(bulletStyles);
assertEqual(bulletResult.bullet, 'Bullet', 'Bullet style should be detected');
assertEqual(bulletResult.numbered, null, 'Numbered should be null');

// Test alternative bullet names
const bulletedStyles = {
    'bulleted': 'Bulleted'
};
assertEqual(detectListStyles(bulletedStyles).bullet, 'Bulleted', 'Bulleted fallback should work');

const bulletedListStyles = {
    'bulleted list': 'Bulleted List'
};
assertEqual(detectListStyles(bulletedListStyles).bullet, 'Bulleted List', 'Bulleted List fallback should work');

// Test numbered style detection
const numberedStyles = {
    'numbered': 'Numbered'
};
const numberedResult = detectListStyles(numberedStyles);
assertEqual(numberedResult.numbered, 'Numbered', 'Numbered style should be detected');
assertEqual(numberedResult.bullet, null, 'Bullet should be null');

// Test numbered list fallback
const numberedListStyles = {
    'numbered list': 'Numbered List'
};
assertEqual(detectListStyles(numberedListStyles).numbered, 'Numbered List', 'Numbered List fallback should work');

// Test preference order (bullet first)
const allBulletStyles = {
    'bullet': 'Bullet',
    'bulleted': 'Bulleted',
    'bulleted list': 'Bulleted List'
};
assertEqual(detectListStyles(allBulletStyles).bullet, 'Bullet', 'Bullet should take precedence');

// Test both list styles together
const bothListStyles = {
    'bullet': 'Bullet',
    'numbered': 'Numbered'
};
const bothResult = detectListStyles(bothListStyles);
assertEqual(bothResult.bullet, 'Bullet', 'Bullet found with both');
assertEqual(bothResult.numbered, 'Numbered', 'Numbered found with both');

// =============================================================================
// SUMMARY
// =============================================================================

console.log('\n' + '='.repeat(50));
console.log(`Results: ${passed} passed, ${failed} failed`);

if (failed > 0) {
    process.exit(1);
} else {
    console.log('All tests passed!');
    process.exit(0);
}
