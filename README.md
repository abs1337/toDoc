<a name="toDoc"></a>

## toDoc
JavaScript library to convert HTML/Text to a Word Documnet


### toDoc.createSection(sectionType, contentType, content, alignment, position, nextLine)
Creates a Header or Footer Section in the document

**Kind**: static method of [<code>toDoc</code>](#toDoc)
**Access**: public

| Param | Type | Description |
| --- | --- | --- |
| sectionType | <code>string</code> | Defines whether the section is a header or footer <br/> Accepts : "header", "footer" <br/> Required |
| contentType | <code>string</code> | Defines whether the section contains an image, text or HTML markup <br/> Accepts : "image", "text", "html" <br/> Required |
| content | <code>string</code> | Defines the sections's content <br/> Accepts : stringified image URLs, stringified text, stringified HTML markup <br/> Required |
| alignment | <code>string</code> | Defines the section content's alignemnt <br/> Accepts : "left", "center", "right" <br/> Required |
| position | <code>number</code> | Specifies the content's position in the section <br/> Accepts : 1++ <br/> Default value :  0 <br/> Optional |
| nextLine | <code>boolean</code> | Specifes whether the section's content should start in a new line <br/> Accepts : true, false <br/> Default value : false <br/> Optional |

<a name="toDoc.doc.createContent"></a>

### toDoc.createContent(type, content, position, nextLine)
Creates a Paragraph or Page in the document

**Kind**: static method of [<code>toDoc</code>](#toDoc)
**Access**: public

| Param | Type | Description |
| --- | --- | --- |
| type | <code>string</code> | Specify whether the content is a Paragraph, a Page or an Image <br/> Accepts : "paragraph", "page", "image" <br/> Required |
| content | <code>string</code> | Defines the document's content <br/> Accepts : stringified text, stringified HTML markup, stringified image URLs <br/> Required |
| position | <code>number</code> | Defines the content's position in the document <br/> Accepts : 1++ <br/> Default value : 0 <br/> Required for pagagraphs, pages and images <br/> Optional if passing a whole HTML document |
| nextLine | <code>boolean</code> | Specifes whether the content should start in a new line <br/> Accepts : true, false <br/> Default value : false <br/> Optional |

<a name="toDoc.doc.createDocument"></a>

### toDoc.createDocument(fileName, params)
Generates and saves a Word document.

**Kind**: static method of [<code>toDoc</code>](#toDoc)
**Access**: public

| Param | Type | Description |
| --- | --- | --- |
| fileName | <code>string</code> | Specifies the name of the document <br/> Required |
| params | <code>object</code> | Define custom parameters for the document's layout <br/> Optional |
