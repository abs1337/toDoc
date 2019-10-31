## toDoc : <code>object</code>
Convert HTML/Text to a Word Documnet

**Kind**: global namespace

* [toDoc](#toDoc) : <code>object</code>
    * [.toDoc()](#toDoc.toDoc) ⇒ <code>object</code>
    * [.doc.createSection(sectionType, contentType, content, position, align, size, font)](#toDoc.doc.createSection)
    * [.doc.createContent(type, content, position, align, string, font)](#toDoc.doc.createContent)
    * [.doc.createImage(sectionType, imageURL, position, align, imageWidth, imageHeight)](#toDoc.doc.createImage)
    * [.doc.createPagenum(sectionType, format, align, string, font)](#toDoc.doc.createPagenum)
    * [.doc.clearDocument(fileName, params)](#toDoc.doc.clearDocument)
    * [.doc.createDocument(fileName, params)](#toDoc.doc.createDocument)

<a name="toDoc.toDoc"></a>

### toDoc.toDoc() ⇒ <code>object</code>
Master function

**Kind**: static method of [<code>toDoc</code>](#toDoc)
**Returns**: <code>object</code> - Library object with global functions
**Access**: public
<a name="toDoc.doc.createSection"></a>

### toDoc.doc.createSection(sectionType, contentType, content, position, align, size, font)
Creates a Header or a Footer Section in the document

**Kind**: static method of [<code>toDoc</code>](#toDoc)
**Access**: public

| Param | Type | Description |
| --- | --- | --- |
| sectionType | <code>string</code> | Specifies whether the content is for the header or footer of the document <br/> Accepts : "header" &#124; "footer" <br/> Required |
| contentType | <code>string</code> | Defines whether the section's content is Text or stringified HTML markup <br/> Accepts : "text" &#124; "html" <br/> Required |
| content | <code>string</code> | The content that will be created in section <br/> Accepts : stringified text &#124; stringified HTML markup <br/> Required |
| position | <code>number</code> | Specifies the content's position in the section <br/> Accepts : 1++ <br/> Default position: 0 <br/> Optional |
| align | <code>string</code> | Defines the section content's alignemnt <br/> Accepts : "left" &#124; "center" &#124; "right" &#124; "justify" <br/> Default alignment: left <br/> Optional |
| size | <code>string</code> | Specifies the content's font size <br/> Accepts : CSS font-size <br/> Default size: 12pt/16px/1em  <br/> Optional |
| font | <code>string</code> | Specifies the content's font family <br/> Accepts : CSS font-family <br/> Default font: Times New Roman <br/> Optional |

<a name="toDoc.doc.createContent"></a>

### toDoc.doc.createContent(type, content, position, align, string, font)
Creates a Paragraph or Page in the document

**Kind**: static method of [<code>toDoc</code>](#toDoc)
**Access**: public

| Param | Type | Description |
| --- | --- | --- |
| type | <code>string</code> | Specify whether the content is a Paragraph or a new Page <br/> Accepts : "paragraph" &#124; "page" <br/> Required |
| content | <code>string</code> | The content that will be created as a Paragraph or Page <br/> Accepts : stringified text &#124; stringified HTML markup <br/> Required |
| position | <code>number</code> | Defines the content's position in the document <br/> Accepts : 1++ <br/> Default value : 0 <br/> Required for pagagraphs and pages <br/> Optional if passing a whole HTML document |
| align | <code>string</code> | Defines the content's alignemnt <br/> Accepts : "left" &#124; "center" &#124; "right" &#124; "justify" <br/> Default alignment: left <br/> Optional |
| string | <code>string</code> | Specifies the content's font size <br/> Accepts : CSS font-size <br/> Default size: 12pt/16px/1em  <br/> Optional |
| font | <code>string</code> | Specifies the content's font family <br/> Accepts : CSS font-family <br/> Default font: Times New Roman <br/> Optional |

<a name="toDoc.doc.createImage"></a>

### toDoc.doc.createImage(sectionType, imageURL, position, align, imageWidth, imageHeight)
Creates an image in the Header, Footer or Body  of the document

**Kind**: static method of [<code>toDoc</code>](#toDoc)
**Access**: public

| Param | Type | Description |
| --- | --- | --- |
| sectionType | <code>string</code> | Specifies where the image will be inserted in the document <br/> Accepts : "header" &#124; "footer" &#124; "body" <br/> Required |
| imageURL | <code>string</code> | Specifies the image's URL <br/> Accepts : stringified image URL <br/> Required |
| position | <code>number</code> | Specifies the image's position in the document <br/> Accepts : 1++ <br/> Default value: 0 <br/> Optional |
| align | <code>string</code> | Defines the image's alignemnt <br/> Accepts : "left" &#124; "center" &#124; "right" <br/> Default alignment: left <br/> Optional |
| imageWidth | <code>number</code> | Specify a custom width for images <br/> Only works for type: "image" <br/> Optional |
| imageHeight | <code>number</code> | Specify a custom height for images <br/> Only works for type: "image" <br/> Optional |

<a name="toDoc.doc.createPagenum"></a>

### toDoc.doc.createPagenum(sectionType, format, align, string, font)
Inserts page number in specified section of the document

**Kind**: static method of [<code>toDoc</code>](#toDoc)
**Access**: public

| Param | Type | Description |
| --- | --- | --- |
| sectionType | <code>string</code> | Specifies where the page number is inserted in the document <br/> Accepts : "header" &#124; "footer" <br/> Required |
| format | <code>number</code> | Specifies the page number's format <br/> Accepts : 1 &#124; 2 <br/> Default value: 1 <br/> Optional |
| align | <code>string</code> | Defines the page number's alignemnt <br/> Accepts : "left" &#124; "center" &#124; "right" <br/> Default alignment: left <br/> Optional |
| string | <code>string</code> | Specifies the page number's font size <br/> Accepts : CSS font-size <br/> Default size: 12pt/16px/1em  <br/> Optional |
| font | <code>string</code> | Specifies the page number's font family <br/> Accepts : CSS font-family <br/> Default font: Times New Roman <br/> Optional |

<a name="toDoc.doc.clearDocument"></a>

### toDoc.doc.clearDocument(fileName, params)
Clears all section and content data

**Kind**: static method of [<code>toDoc</code>](#toDoc)
**Access**: public

| Param | Type | Description |
| --- | --- | --- |
| fileName | <code>string</code> | Specifies the name of the document <br/> Required |
| params | <code>object</code> | Define custom parameters for the document's layout <br/> Optional |

<a name="toDoc.doc.createDocument"></a>

### toDoc.doc.createDocument(fileName, params)
Generates and saves a Word document.

**Kind**: static method of [<code>toDoc</code>](#toDoc)
**Access**: public

| Param | Type | Description |
| --- | --- | --- |
| fileName | <code>string</code> | Specifies the name of the document <br/> Required |
| params | <code>object</code> | Define custom parameters for the document's layout <br/> Optional |
