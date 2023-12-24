document.addEventListener("DOMContentLoaded", function() {
    document.getElementById("run").addEventListener("click", run);
});

async function run() {
    try {
        await Word.run(async context => {
            const originalText = "The quick brown fox jumps over the lazy dog.";
            const revisedText = "The fast brown cat jumps over the lazy dog and car.";

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
        });
    } catch (error) {
        console.error("Error: " + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.error("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }
}

function getDiff(originalText, revisedText) {
    const diff = JsDiff.diffWords(originalText, revisedText);
    return diff;
}
