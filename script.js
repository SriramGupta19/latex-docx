async function convertToDocx() {
    const latex = document.getElementById('latexInput').value;
    const zip = new JSZip();

    // 1. Create the File Structure (Required for Word)
    const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
        </Relationships>`;
    
    const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
            <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
            <Default Extension="xml" ContentType="application/xml"/>
            <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
        </Types>`;

    // 2. Parse LaTeX and build document.xml
    const docXml = buildDocumentXml(latex);

    // 3. Assemble the Zip
    zip.file("_rels/.rels", rels);
    zip.file("[Content_Types].xml", contentTypes);
    zip.file("word/document.xml", docXml);

    // 4. Generate and Download
    const content = await zip.generateAsync({type: "blob"});
    const link = document.createElement('a');
    link.href = URL.createObjectURL(content);
    link.download = "Thesis_Converted.docx";
    link.click();
}

function buildDocumentXml(latex) {
    let bodyContent = "";

    // Basic Regex Parsers for your Thesis structure
    const lines = latex.split('\n');

    lines.forEach(line => {
        // Handle Sections
        if (line.match(/\\section\{(.+?)\}/)) {
            const title = line.match(/\\section\{(.+?)\}/)[1];
            bodyContent += `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>${title}</w:t></w:r></w:p>`;
        } 
        // Handle Subsection (e.g., 2.4.1 APF Overview) [cite: 49]
        else if (line.match(/\\subsection\{(.+?)\}/)) {
            const sub = line.match(/\\subsection\{(.+?)\}/)[1];
            bodyContent += `<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>${sub}</w:t></w:r></w:p>`;
        }
        // Handle Plain Text & Equations
        else if (line.trim() !== "" && !line.startsWith('\\')) {
            // Basic sanitization for XML
            const cleanText = line.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
            bodyContent += `<w:p><w:r><w:t>${cleanText}</w:t></w:r></w:p>`;
        }
    });

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:body>
                ${bodyContent}
                <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
            </w:body>
        </w:document>`;
}
