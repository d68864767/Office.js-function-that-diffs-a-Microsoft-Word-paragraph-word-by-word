// test.js

const assert = require('assert');
const { getDiff, updateParagraph } = require('./word.js');

describe('Word Paragraph Update', function() {
    describe('#getDiff()', function() {
        it('should return correct diff for given texts', function() {
            const originalText = "The quick brown fox jumps over the lazy dog.";
            const revisedText = "The fast brown cat jumps over the lazy dog and car.";
            const expectedDiff = [
                { value: 'The ', count: 1 },
                { value: 'quick', count: 1, removed: true },
                { value: 'fast', count: 1, added: true },
                { value: ' brown ', count: 1 },
                { value: 'fox', count: 1, removed: true },
                { value: 'cat', count: 1, added: true },
                { value: ' jumps over the lazy dog', count: 1 },
                { value: ' and car', count: 1, added: true }
            ];
            const diff = getDiff(originalText, revisedText);
            assert.deepStrictEqual(diff, expectedDiff);
        });
    });

    describe('#updateParagraph()', function() {
        it('should update the paragraph correctly', async function() {
            const originalText = "The quick brown fox jumps over the lazy dog.";
            const revisedText = "The fast brown cat jumps over the lazy dog and car.";
            const context = {
                document: {
                    getSelection: function() {
                        return {
                            load: function() {},
                            paragraphs: {
                                getFirst: function() {
                                    return {
                                        load: function() {},
                                        text: originalText,
                                        getRange: function() {
                                            return {
                                                select: function() {},
                                                insertText: function() {}
                                            };
                                        }
                                    };
                                }
                            }
                        };
                    },
                    sync: function() {
                        return new Promise(resolve => resolve());
                    }
                }
            };
            await updateParagraph(context, originalText, revisedText);
        });
    });
});
