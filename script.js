async function convertToDocx() {
    const latex = document.getElementById('latexInput').value;
    const zip = new JSZip();

    // 1. Content Types: Defines what files are in the zip
    const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
        <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
        <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
    </Types>`;

    // 2. Relationships: Tells Word where the document content is
    const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
    </Relationships>`;

    // 3. Document Content
    const docXml = buildDocumentXml(latex);

    // 4. Basic Styles (Crucial for Headings to look right)
    const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:style w:type="paragraph" w:styleId="Heading1">
            <w:name w:val="heading 1"/><w:rPr><w:b/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr>
        </w:style>
        <w:style w:type="paragraph" w:styleId="Heading2">
            <w:name w:val="heading 2"/><w:rPr><w:b/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>
        </w:style>
    </w:styles>`;

    zip.file("[Content_Types].xml", contentTypes);
    zip.file("_rels/.rels", rels);
    zip.file("word/document.xml", docXml);
    zip.file("word/styles.xml", stylesXml);

    const content = await zip.generateAsync({type: "blob"});
    const link = document.createElement('a');
    link.href = URL.createObjectURL(content);
    link.download = "Converted_Thesis.docx";
    link.click();
}

function buildDocumentXml(latex) {
    let body = "";
    
    // Split into paragraphs by double newlines
    const paragraphs = latex.split(/\n\s*\n/);

    paragraphs.forEach(p => {
        let trimP = p.trim();
        if (!trimP) return;

        // Check for Headings
        if (trimP.startsWith('\\section')) {
            const text = trimP.match(/\\section\{(.+?)\}/)?.[1] || "";
            body += `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
        } 
        else if (trimP.startsWith('\\subsection')) {
            const text = trimP.match(/\\subsection\{(.+?)\}/)?.[1] || "";
            body += `<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
        } 
        else {
            // Regular Paragraph with Inline Logic (Math, Bold, Italics)
            body += `<w:p>${processInlineFormatting(trimP)}</w:p>`;
        }
    });

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" 
                xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
        <w:body>${body}</w:body>
    </w:document>`;
}

function processInlineFormatting(text) {
    let result = "";
    // Regex to split text by math delimiters $...$
    const parts = text.split(/(\$.+?\$)/g);

    parts.forEach(part => {
        if (part.startsWith('$') && part.endsWith('$')) {
            // It's Math
            const mathContent = part.substring(1, part.length - 1);
            result += renderOMML(mathContent);
        } else {
            // It's Text - Handle Bold/Italic
            let processedText = part
                .replace(/\\textbf\{(.+?)\}/g, '<w:r><w:rPr><w:b/></w:rPr><w:t>$1</w:t></w:r>')
                .replace(/\\textit\{(.+?)\}/g, '<w:r><w:rPr><w:i/></w:rPr><w:t>$1</w:t></w:r>');
            
            // If no LaTeX commands were found in this chunk, wrap it in a standard run
            if (!processedText.includes('<w:r>')) {
                result += `<w:r><w:t xml:space="preserve">${processedText}</w:t></w:r>`;
            } else {
                result += processedText;
            }
        }
    });
    return result;
}

function renderOMML(latex) {
    let xml = latex;

    // 1. Handle Fractions: \frac{a}{b}
    xml = xml.replace(/\\frac\{(.+?)\}\{(.+?)\}/g, 
        `<m:f><m:num><m:r><m:t>$1</m:t></m:r></m:num><m:den><m:r><m:t>$2</m:t></m:r></m:den></m:f>`);

    // 2. Handle Subscripts: x_i or x_{abc}
    xml = xml.replace(/([a-zA-Z0-9])_\{?([a-zA-Z0-9]+)\}?/g, 
        `<m:sSub><m:e><m:r><m:t>$1</m:t></m:r></m:e><m:sub><m:r><m:t>$2</m:t></m:r></m:sub></m:sSub>`);

    // 3. Handle Superscripts: x^2 or x^{10}
    xml = xml.replace(/([a-zA-Z0-9])\^\{?([a-zA-Z0-9]+)\}?/g, 
        `<m:sSup><m:e><m:r><m:t>$1</m:t></m:r></m:e><m:sup><m:r><m:t>$2</m:t></m:r></m:sup></m:sSup>`);

    // 4. Basic Greek/Symbols Replacement
    const symbols = {
        '\\theta': 'θ', '\\lambda': 'λ', '\\omega': 'ω', '\\sigma': 'σ',
        '\\alpha': 'α', '\\beta': 'β', '\\gamma': 'γ', '\\Delta': 'Δ',
        '\\pi': 'π', '\\infty': '∞', '\\pm': '±', '\\times': '×'
    };
    Object.keys(symbols).forEach(key => {
        xml = xml.replace(new RegExp('\\' + key, 'g'), symbols[key]);
    });

    // 5. Wrap leftover plain characters in math runs
    // This finds sequences of chars that aren't inside existing <m: tags
    if (!xml.includes('<m:')) {
        xml = `<m:r><m:t>${xml}</m:t></m:r>`;
    }

    return `<m:oMath>${xml}</m:oMath>`;
}
