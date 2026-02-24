async function convertToDocx() {
    const latex = document.getElementById('latexInput').value;
    const zip = new JSZip();

    // Word Content Types and Relations definitions
    const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`;
    const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`;

    // Core LaTeX to Word XML Conversion
    const docXml = buildDocumentXml(latex);

    zip.file("[Content_Types].xml", contentTypes);
    zip.file("_rels/.rels", rels);
    zip.file("word/document.xml", docXml);

    const content = await zip.generateAsync({type: "blob"});
    const link = document.createElement('a');
    link.href = URL.createObjectURL(content);
    link.download = "Thesis_Draft.docx";
    link.click();
}

function buildDocumentXml(latex) {
    let body = "";
    const lines = latex.split('\n');

    lines.forEach(line => {
        // Match Section Headers
        if (line.match(/\\section\{(.+?)\}/)) {
            const txt = line.match(/\\section\{(.+?)\}/)[1];
            body += `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>${txt}</w:t></w:r></w:p>`;
        } 
        // Match Formulas (Inline $...$ or Block $$...$$)
        else if (line.includes('$')) {
            body += `<w:p>`;
            let parts = line.split('$');
            parts.forEach((part, index) => {
                if (index % 2 === 0) { // Plain text
                    body += `<w:r><w:t xml:space="preserve">${part}</w:t></w:r>`;
                } else { // Math block
                    body += renderOMML(part);
                }
            });
            body += `</w:p>`;
        }
        // Plain Paragraphs
        else if (line.trim() !== "") {
            body += `<w:p><w:r><w:t>${line}</w:t></w:r></w:p>`;
        }
    });

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
        <w:body>${body}</w:body>
    </w:document>`;
}

function renderOMML(latex) {
    // Basic mapping for thesis variables (e.g., subscripts k_att, lambda_e)
    let math = latex
        .replace(/([a-zA-Z0-9])_([a-zA-Z0-9])/g, '</m:t></m:r><m:sSub><m:e><m:r><m:t>$1</m:t></m:r></m:e><m:num><m:r><m:t>$2</m:t></m:r></m:num></m:sSub><m:r><m:t>')
        .replace(/\\frac\{(.+?)\}\{(.+?)\}/g, '</m:t></m:r><m:f><m:num><m:r><m:t>$1</m:t></m:r></m:num><m:den><m:r><m:t>$2</m:t></m:r></m:den></m:f><m:r><m:t>')
        .replace(/\\theta/g, 'θ').replace(/\\lambda/g, 'λ').replace(/\\omega/g, 'ω').replace(/\\dot\{e\}/g, 'ė');

    return `<m:oMath><m:r><m:t>${math}</m:t></m:r></m:oMath>`;
}
