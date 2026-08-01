import assert from "node:assert/strict";
import test from "node:test";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { addPart, addRelationship, addResourcePart, deleteDocxParagraph, deleteDocxRun, deletePptxShape, duplicatePptxShape, exportDocx, exportPptx, garbageCollectRelationships, importDocx, importPptx, insertDocxParagraph, insertDocxRun, loadOpcPackage, repackOfficePackage, resolveStyle, splitDocxRun, unpackOfficePackage, validateOpcPackage, writeOpcPackage, addPptxShape } from "../src/index.js";

const zip = (entries: Record<string, string | Uint8Array>) => zipSync(Object.fromEntries(Object.entries(entries).map(([name, value]) => [name, typeof value === "string" ? strToU8(value) : value])));
const contentTypes = (overrides: string) => `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="bin" ContentType="application/octet-stream"/>${overrides}</Types>`;
const rels = (items: string) => `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items}</Relationships>`;

function docxFixture(): Uint8Array {
  return zip({
    "[Content_Types].xml": contentTypes('<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'),
    "_rels/.rels": rels('<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'),
    "word/document.xml": '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><w:body><w:p><w:pPr><w:pStyle w:val="Normal"/><w:jc w:val="left"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="24"/></w:rPr><w:t>Hello DOCX</w:t></w:r><w:r><w:drawing><a:blip r:embed="rIdImage"/></w:drawing></w:r></w:p><w:tbl><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr><w:headerReference r:id="rIdHeader"/><w:footerReference r:id="rIdFooter"/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>',
    "word/_rels/document.xml.rels": rels('<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rIdNumbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/><Relationship Id="rIdHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rIdFooter" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/><Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>'),
    "word/styles.xml": '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>',
    "word/numbering.xml": '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="0"/></w:numbering>',
    "word/header1.xml": '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:hdr>',
    "word/footer1.xml": '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Footer</w:t></w:r></w:p></w:ftr>',
    "word/media/image1.png": new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]),
    "customXml/item1.xml": '<x:data xmlns:x="urn:vendor"><x:future keep="yes"/></x:data>',
    "custom/opaque.bin": new Uint8Array([0, 255, 17, 34, 51]),
  });
}

function pptxFixture(): Uint8Array {
  return zip({
    "[Content_Types].xml": contentTypes('<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'),
    "_rels/.rels": rels('<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>'),
    "ppt/presentation.xml": '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldMasterIdLst><p:sldMasterId r:id="rMaster"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id="256" r:id="rSlide"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>',
    "ppt/_rels/presentation.xml.rels": rels('<Relationship Id="rMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/><Relationship Id="rSlide" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>'),
    "ppt/slides/slide1.xml": '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr><p:txBody><a:bodyPr/><a:p><a:r><a:rPr b="1"/><a:t>Hello PPTX</a:t></a:r></a:p></p:txBody></p:sp><p:pic><p:nvPicPr><p:cNvPr id="3" name="Image"/></p:nvPicPr><p:blipFill><a:blip r:embed="rImage"/></p:blipFill><p:spPr/></p:pic><p:grpSp><p:nvGrpSpPr><p:cNvPr id="4" name="Group"/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="5" name="Grouped"/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:t>Group text</a:t></a:r></a:p></p:txBody></p:sp></p:grpSp></p:spTree></p:cSld></p:sld>',
    "ppt/slides/_rels/slide1.xml.rels": rels('<Relationship Id="rLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>'),
    "ppt/slideMasters/slideMaster1.xml": '<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree/></p:cSld><p:sldLayoutIdLst><p:sldLayoutId r:id="rLayout"/></p:sldLayoutIdLst></p:sldMaster>',
    "ppt/slideMasters/_rels/slideMaster1.xml.rels": rels('<Relationship Id="rLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>'),
    "ppt/slideLayouts/slideLayout1.xml": '<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree/></p:cSld></p:sldLayout>',
    "ppt/slideLayouts/_rels/slideLayout1.xml.rels": rels('<Relationship Id="rMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>'),
    "ppt/theme/theme1.xml": '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office"><a:themeElements><a:clrScheme name="Office"><a:accent1><a:srgbClr val="4472C4"/></a:accent1></a:clrScheme></a:themeElements></a:theme>',
    "ppt/media/image1.png": new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]),
  });
}

test("Phase 2A: OPC package preserves XML and binary parts across repack", () => {
  const input = docxFixture(); const original = unzipSync(input); const output = repackOfficePackage(unpackOfficePackage(input)); const repacked = unzipSync(output);
  assert.deepEqual(Object.keys(repacked).sort(), Object.keys(original).sort());
  assert.deepEqual(repacked["word/media/image1.png"], original["word/media/image1.png"]);
  assert.deepEqual(validateOpcPackage(loadOpcPackage(output)), []);
});

test("Phase 2B: DOCX imports core semantics and writes edited body/header text", () => {
  const document = importDocx(docxFixture()); const paragraph = document.root.sections[0]!.blocks[0]!;
  assert.equal(paragraph.type, "paragraph"); if (paragraph.type !== "paragraph" || !("runs" in paragraph)) throw new Error("paragraph expected");
  assert.equal(paragraph.runs.some((run) => run.type === "inlineImage"), true); assert.equal(Object.keys(document.styles.styles).includes("Normal"), true);
  const text = paragraph.runs.find((run) => run.type === "textRun"); if (!text || text.type !== "textRun") throw new Error("text expected"); text.text = "Edited DOCX";
  const headerText = document.root.headers?.[0]?.blocks[0]; if (headerText?.type === "paragraph" && "runs" in headerText && headerText.runs[0]?.type === "textRun") headerText.runs[0].text = "Edited Header";
  const output = exportDocx(document); const entries = unzipSync(output); assert.match(strFromU8(entries["word/document.xml"]!), /Edited DOCX/); assert.match(strFromU8(entries["word/header1.xml"]!), /Edited Header/); assert.deepEqual(validateOpcPackage(loadOpcPackage(output)), []);
});

test("Phase 2C: PPTX imports inheritance/shape tree and writes edited shape text", () => {
  const document = importPptx(pptxFixture()); assert.equal(document.root.slideMasters.length, 1); assert.equal(document.root.slideLayouts.length, 1); assert.equal(Object.keys(document.themes.themes).length, 1);
  const slide = document.root.slides[0]!; assert.equal(slide.layoutRef, "/ppt/slideLayouts/slideLayout1.xml"); assert.equal(slide.shapes.some((shape) => shape.type === "picture"), true); assert.equal(slide.shapes.some((shape) => shape.type === "group"), true);
  const shape = slide.shapes.find((item) => item.type === "placeholder"); if (!shape || shape.type !== "placeholder" || !shape.textBody) throw new Error("placeholder expected"); const text = shape.textBody.paragraphs[0]?.runs[0]; if (!text || text.type !== "textRun") throw new Error("text expected"); text.text = "Edited PPTX";
  const output = exportPptx(document); const entries = unzipSync(output); assert.match(strFromU8(entries["ppt/slides/slide1.xml"]!), /Edited PPTX/); assert.deepEqual(validateOpcPackage(loadOpcPackage(output)), []);
});

test("Phase 3: DOCX run and paragraph structural mutations generate valid XML patches", () => {
  const splitDocument = importDocx(docxFixture()); const splitParagraph = splitDocument.root.sections[0]!.blocks[0]!; if (splitParagraph.type !== "paragraph" || !("runs" in splitParagraph) || splitParagraph.runs[0]?.type !== "textRun") throw new Error("paragraph expected"); splitDocxRun(splitDocument, splitParagraph, splitParagraph.runs[0], 5); const splitXml = strFromU8(unzipSync(exportDocx(splitDocument))["word/document.xml"]!); assert.match(splitXml, /Hello<\/w:t><\/w:r><w:r>/); assert.match(splitXml, / DOCX/);
  const insertDocument = importDocx(docxFixture()); const paragraph = insertDocument.root.sections[0]!.blocks[0]!; if (paragraph.type !== "paragraph" || !("runs" in paragraph)) throw new Error("paragraph expected"); const inserted = insertDocxRun(insertDocument, paragraph, 1, " INSERTED ", { italic: true }); const insertedXml = strFromU8(unzipSync(exportDocx(insertDocument))["word/document.xml"]!); assert.match(insertedXml, /INSERTED/); assert.match(insertedXml, /<w:i\/>/); assert.equal(inserted.source, undefined);
  const deleteRunDocument = importDocx(docxFixture()); const deleteRunParagraph = deleteRunDocument.root.sections[0]!.blocks[0]!; if (deleteRunParagraph.type !== "paragraph" || !("runs" in deleteRunParagraph) || deleteRunParagraph.runs[0]?.type !== "textRun") throw new Error("paragraph expected"); deleteDocxRun(deleteRunDocument, deleteRunParagraph, deleteRunParagraph.runs[0]); assert.doesNotMatch(strFromU8(unzipSync(exportDocx(deleteRunDocument))["word/document.xml"]!), /Hello DOCX/);
  const paragraphDocument = importDocx(docxFixture()); const section = paragraphDocument.root.sections[0]!; insertDocxParagraph(paragraphDocument, section, 1, "New paragraph"); assert.match(strFromU8(unzipSync(exportDocx(paragraphDocument))["word/document.xml"]!), /New paragraph/);
  const deleteParagraphDocument = importDocx(docxFixture()); const deleteSection = deleteParagraphDocument.root.sections[0]!; const originalParagraph = deleteSection.blocks[0]!; if (originalParagraph.type !== "paragraph" || !("runs" in originalParagraph)) throw new Error("paragraph expected"); deleteDocxParagraph(deleteParagraphDocument, deleteSection, originalParagraph); assert.doesNotMatch(strFromU8(unzipSync(exportDocx(deleteParagraphDocument))["word/document.xml"]!), /Hello DOCX/);
});

test("Phase 3: PPTX shape add/delete/duplicate patches the shape tree", () => {
  const addDocument = importPptx(pptxFixture()); addPptxShape(addDocument, addDocument.root.slides[0]!, "New shape"); assert.match(strFromU8(unzipSync(exportPptx(addDocument))["ppt/slides/slide1.xml"]!), /New shape/);
  const deleteDocument = importPptx(pptxFixture()); const deleteSlide = deleteDocument.root.slides[0]!; deletePptxShape(deleteDocument, deleteSlide, deleteSlide.shapes[0]!); assert.doesNotMatch(strFromU8(unzipSync(exportPptx(deleteDocument))["ppt/slides/slide1.xml"]!), /Hello PPTX/);
  const duplicateDocument = importPptx(pptxFixture()); const duplicateSlide = duplicateDocument.root.slides[0]!; duplicatePptxShape(duplicateDocument, duplicateSlide, duplicateSlide.shapes[0]!); const duplicatedXml = strFromU8(unzipSync(exportPptx(duplicateDocument))["ppt/slides/slide1.xml"]!); assert.equal(duplicatedXml.match(/Hello PPTX/g)?.length, 2);
});

test("Phase 3: relationship/resource lifecycle and content types stay consistent", () => {
  const pkg = loadOpcPackage(docxFixture()); const documentUri = "/word/document.xml";
  const added = addResourcePart(pkg, { sourcePartUri: documentUri, directory: "/word/media", fileName: "image1.png", contentType: "image/png", relationshipType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", data: new Uint8Array([1, 2, 3]) });
  assert.equal(added.part.uri, "/word/media/image12.png"); assert.equal(pkg.parts.has(added.part.uri), true); assert.equal(pkg.contentTypes.defaults.png, "image/png");
  const unused = addRelationship(pkg, documentUri, "urn:test:unused", "unused.xml"); addPart(pkg, { preferredUri: "/word/unused.xml", contentType: "application/xml", data: "<unused/>" }); const removed = garbageCollectRelationships(pkg, documentUri); assert.equal(removed.some((item) => item.id === unused.id), true);
  assert.deepEqual(validateOpcPackage(loadOpcPackage(writeOpcPackage(pkg))), []);
});

test("Phase 3: style resolution follows defaults, inheritance, theme, and direct formatting", () => {
  const registry = { documentDefaults: { fontSize: 10, color: "black" }, styles: { Base: { id: "Base", kind: "paragraph" as const, properties: { fontFamily: "Aptos", color: "blue" } }, Heading: { id: "Heading", kind: "paragraph" as const, basedOn: "Base", properties: { fontSize: 18, bold: true } } } };
  assert.deepEqual(resolveStyle({ registry, styleId: "Heading", themeDefaults: { color: "green" }, direct: { bold: false } }), { fontSize: 18, color: "blue", fontFamily: "Aptos", bold: false });
});

test("Phase 3: complex unknown XML and opaque binary bytes survive semantic editing", () => {
  const input = unzipSync(docxFixture()); const document = importDocx(docxFixture()); const paragraph = document.root.sections[0]!.blocks[0]!; if (paragraph.type !== "paragraph" || !("runs" in paragraph) || paragraph.runs[0]?.type !== "textRun") throw new Error("paragraph expected"); paragraph.runs[0].text = "Preservation edit"; const output = unzipSync(exportDocx(document)); assert.deepEqual(output["customXml/item1.xml"], input["customXml/item1.xml"]); assert.deepEqual(output["custom/opaque.bin"], input["custom/opaque.bin"]);
});
