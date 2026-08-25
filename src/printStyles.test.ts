import fs from 'fs';
import path from 'path';

/**
 * 元件測試跑在 jsdom 上、不會套用樣式表，因此抓不到 CSS 串接順序的問題。
 * 這份測試直接讀 index.css，鎖住列印樣式裡「隱藏必須勝過版面」的順序不變式。
 *
 * 背景：`.print\:hidden` 曾經寫在 `.grid-cols-1 { display: block !important }` 前面，
 * 兩者同權重同 !important，由後者勝出，導致帶 grid-cols-1 的區塊即使標了 print:hidden
 * 仍會出現在匯出的 PDF 裡。
 */
const readPrintBlock = (): string => {
    const css = fs.readFileSync(path.join(__dirname, 'index.css'), 'utf8');
    const start = css.indexOf('@media print');
    expect(start).toBeGreaterThan(-1);

    let depth = 0;
    for (let i = css.indexOf('{', start); i < css.length; i++) {
        if (css[i] === '{') depth++;
        if (css[i] === '}') {
            depth--;
            if (depth === 0) return css.slice(start, i + 1);
        }
    }

    throw new Error('找不到完整的 @media print 區塊');
};

describe('print stylesheet cascade', () => {
    const printBlock = readPrintBlock();

    it('declares the hide rules after every !important display rule', () => {
        const hideRuleIndex = printBlock.indexOf('.print\\:hidden');
        expect(hideRuleIndex).toBeGreaterThan(-1);

        const overridePattern = /display:\s*(block|grid)\s*!important/g;
        const overrideIndexes: number[] = [];
        let match = overridePattern.exec(printBlock);
        while (match !== null) {
            overrideIndexes.push(match.index);
            match = overridePattern.exec(printBlock);
        }

        expect(overrideIndexes.length).toBeGreaterThan(0);
        overrideIndexes.forEach((index) => {
            expect(index).toBeLessThan(hideRuleIndex);
        });
    });

    it('keeps the hide rules marked !important so they beat layout utilities', () => {
        const hideRule = printBlock.slice(printBlock.indexOf('.print\\:hidden'));

        expect(hideRule).toMatch(/display:\s*none\s*!important/);
    });

    it('preserves background colours so the blue report headers survive printing', () => {
        expect(printBlock).toContain('.report-brand-header');
        expect(printBlock).toMatch(/print-color-adjust:\s*exact\s*!important/);
    });
});
