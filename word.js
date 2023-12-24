// word.js

const JsDiff = require('diff');

function getDiff(originalText, revisedText) {
    const diff = JsDiff.diffWords(originalText, revisedText);
    return diff;
}

async function updateParagraph(context, originalText, revisedText) {
    const diff = getDiff(originalText, revisedText);

    const range = context.document.getSelection();
    range.load("text");

    await context.sync();

    const paragraph = range.paragraphs.getFirst();
    paragraph.load("text");

    await context.sync();

    const originalParagraphText = paragraph.text;

    if (originalParagraphText !== originalText) {
        console.error("Selected paragraph does not match the original text.");
        return;
    }

    diff.forEach(change => {
        if (change.removed) {
            const index = originalParagraphText.indexOf(change.value);
            const range = paragraph.getRange('Whole');
            range.select('Start', 'End');
            range.insertText(change.value, 'Replace');
        } else if (change.added) {
            const index = revisedText.indexOf(change.value);
            const range = paragraph.getRange('Whole');
            range.select('Start', 'End');
            range.insertText(change.value, 'Replace');
        }
    });

    await context.sync();
}

module.exports = {
    getDiff,
    updateParagraph
};
