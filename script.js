async function convertToDocx() {
    const latex = document.getElementById('latexInput').value;
    const zip = new JSZip();

    // 1. Content Types & Relationships (Standard OOXML)
    const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`;
    const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`;

    // 2. Styles Definition (Heading 1-3 and Bullets)
    const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:rPr><w:b/><w:sz w:val="48"/><w:szCs w:val="48"/></w:rPr></w:style>
        <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:rPr><w:b/><w:sz w:val="36"/><w:szCs w:val="36"/></w:rPr></w:style>
        <w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:rPr><w:b/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:style>
        <w:style w:type="paragraph" w:styleId="ListBullet"><w:name w:val="List Bullet"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:style>
    </w:styles>`;

    // 3. Process the LaTeX
    const docXml = buildDocumentXml(latex);

    zip.file("[Content_Types].xml", contentTypes);
    zip.file("_rels/.rels", rels);
    zip.file("word/document.xml", docXml);
    zip.file("word/styles.xml", stylesXml);

    const content = await zip.generateAsync({type: "blob"});
    const link = document.createElement('a');
    link.href = URL.createObjectURL(content);
    link.download = "Thesis_Output.docx";
    link.click();
}

function buildDocumentXml(latex) {
    let body = "";
    
    // Pre-processing: Remove LaTeX comments and labels
    let cleanLatex = latex.replace(/%.*$/gm, "").replace(/\\label\{.*?\}/g, "");

    // Split by environments or double newlines
    const blocks = cleanLatex.split(/(\\begin\{[a-z\*]+\}|\\end\{[a-z\*]+\}|\n\s*\n)/);

    let inItemize = false;

    blocks.forEach(block => {
        let trimB = block.trim();
        if (!trimB) return;

        // --- Block Level Elements ---
        if (trimB.includes("\\begin{itemize}")) { inItemize = true; return; }
        if (trimB.includes("\\end{itemize}")) { inItemize = false; return; }
        
        // Headings
        if (trimB.startsWith("\\chapter")) {
            body += `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>${processInline(trimB.match(/\\chapter\{(.+?)\}/)?.[1])}</w:p>`;
        }
        else if (trimB.startsWith("\\section")) {
            body += `<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr>${processInline(trimB.match(/\\section\{(.+?)\}/)?.[1])}</w:p>`;
        }
        else if (trimB.startsWith("\\subsubsection")) {
            body += `<w:p><w:pPr><w:pStyle w:val="Heading3"/></w:pPr>${processInline(trimB.match(/\\subsubsection\{(.+?)\}/)?.[1])}</w:p>`;
        }
        // Equations (Display Mode)
        else if (trimB.includes("\\begin{equation}")) {
            const eqContent = trimB.replace(/\\begin\{equation\}|\\end\{equation\}/g, "");
            body += `<w:p><w:pPr><w:jc w:val="center"/></w:pPr>${renderOMML(eqContent)}</w:p>`;
        }
        // Figures (Placeholder)
        else if (trimB.includes("\\begin{figure}")) {
            const caption = trimB.match(/\\caption\{(.+?)\}/)?.[1] || "Figure";
            const imgPath = trimB.match(/\\includegraphics.*?\{(.+?)\}/)?.[1] || "";
            body += `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:i/></w:rPr><w:t>[IMAGE PLACEHOLDER: ${imgPath}]</w:t></w:r></w:p>`;
            body += `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>Fig: ${caption}</w:t></w:r></w:p>`;
        }
        // List Items
        else if (trimB.startsWith("\\item")) {
            const content = trimB.replace("\\item", "");
            body += `<w:p><w:pPr><w:pStyle w:val="ListBullet"/></w:pPr><w:r><w:t>• </w:t></w:r>${processInline(content)}</w:p>`;
        }
        // Regular Text
        else if (!trimB.startsWith("\\") || trimB.startsWith("\\ref")) {
            body += `<w:p>${processInline(trimB)}</w:p>`;
        }
    });

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
        <w:body>${body}</w:body>
    </w:document>`;
}

function processInline(text) {
    if (!text) return "";
    let result = "";
    
    // Split by inline math \( ... \) or $ ... $ or references \ref{...}
    const parts = text.split(/(\\\(.+?\\\))|(\$.+?\$)|(\\ref\{.+?\})/g);

    parts.forEach(part => {
        if (!part) return;
        
        if (part.startsWith('\\(') || part.startsWith('$')) {
            const cleanMath = part.replace(/\\\(|\\\)|(\$)/g, "");
            result += renderOMML(cleanMath);
        } 
        else if (part.startsWith('\\ref')) {
            const refId = part.match(/\{(.+?)\}/)?.[1] || "??";
            result += `<w:r><w:rPr><w:b/></w:rPr><w:t>${refId}</w:t></w:r>`;
        }
        else {
            result += `<w:r><w:t xml:space="preserve">${part}</w:t></w:r>`;
        }
    });
    return result;
}

function renderOMML(latex) {
    let xml = latex.trim();

    // 1. Math Symbols Mapping
    const symbols = {
        '\\nabla': '∇', '\\theta': 'θ', '\\lambda': 'λ', '\\omega': 'ω', 
        '\\sigma': 'σ', '\\alpha': 'α', '\\beta': 'β', '\\gamma': 'γ', 
        '\\Delta': 'Δ', '\\pi': 'π', '\\infty': '∞', '\\pm': '±', 
        '\\times': '×', '\\|': '||', '\\partial': '∂'
    };
    Object.keys(symbols).forEach(key => {
        xml = xml.replace(new RegExp(key.replace('\\', '\\\\'), 'g'), symbols[key]);
    });

    // 2. Fractions \frac{a}{b}
    xml = xml.replace(/\\frac\{(.+?)\}\{(.+?)\}/g, `<m:f><m:num><m:r><m:t>$1</m:t></m:r></m:num><m:den><m:r><m:t>$2</m:t></m:r></m:den></m:f>`);

    // 3. Subscripts/Superscripts (with nested brackets support)
    xml = xml.replace(/([a-zA-Z0-9∇])_\{?([a-zA-Z0-9, \-]+)\}?/g, `<m:sSub><m:e><m:r><m:t>$1</m:t></m:r></m:e><m:sub><m:r><m:t>$2</m:t></m:r></m:sub></m:sSub>`);
    xml = xml.replace(/([a-zA-Z0-9])\^\{?([a-zA-Z0-9, \-]+)\}?/g, `<m:sSup><m:e><m:r><m:t>$1</m:t></m:r></m:e><m:sup><m:r><m:t>$2</m:t></m:r></m:sup></m:sSup>`);

    // 4. Text in Math \text{...}
    xml = xml.replace(/\\text\{(.+?)\}/g, `<m:r><m:rPr><m:nor/></m:rPr><m:t>$1</m:t></m:r>`);

    // 5. Normal text runs for anything not already converted
    if (!xml.includes('<m:')) {
        xml = `<m:r><m:t>${xml}</m:t></m:r>`;
    }

    return `<m:oMath>${xml}</m:oMath>`;
}
