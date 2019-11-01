## toDoc : <code>object</code>
JS Library to create create a Word document with Text/HTML

**Kind**: global namespace

* [toDoc](#toDoc) : <code>object</code>
    * [.toDoc()](#toDoc.toDoc) ⇒ <code>object</code>
    * [.doc.createSection(sectionType, contentType, content, position, align, size, font)](#toDoc.doc.createSection)
    * [.doc.createContent(type, content, position, align, size, font)](#toDoc.doc.createContent)
    * [.doc.createImage(sectionType, imageURL, position, align, imageWidth, imageHeight)](#toDoc.doc.createImage)
    * [.doc.createPagenumber(sectionType, format, align, size, font)](#toDoc.doc.createPagenumber)
    * [.doc.createDocument(fileName, params)](#toDoc.doc.createDocument)
    * [.doc.clearDocument()](#toDoc.doc.clearDocument)

<a name="toDoc.toDoc"></a>

### toDoc.toDoc() ⇒ <code>object</code>
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
| sectionType | <code>string</code> | Specifies whether to create content for the header or footer of the document <br/> Accepts : "header" &#124; "footer" <br/> Required |
| contentType | <code>string</code> | Defines whether the content is Text or stringified HTML markup <br/> Accepts : "text" &#124; "html" <br/> Required |
| content | <code>string</code> | The content that will be inserted in the document <br/> Accepts : stringified text &#124; stringified HTML markup <br/> Required |
| position | <code>number</code> | Specifies the content's position in the section  <br/> Header/Footer content will be sorted by positon during document generation <br/> Accepts : Numbers > 0 <br/> Default position : 0 <br/> Optional |
| align | <code>string</code> | Specifies the content's alignemnt <br/> Accepts : "left" &#124; "center" &#124; "right" &#124; "justify" <br/> Default alignment : left <br/> Optional |
| size | <code>string</code> | Specifies the content's font size <br/> Accepts : CSS font-size <br/> Default size : 12pt/16px/1em  <br/> Optional |
| font | <code>string</code> | Specifies the content's font family <br/> Accepts : CSS font-family <br/> Default font : Times New Roman <br/> Optional |

<a name="toDoc.doc.createContent"></a>

### toDoc.doc.createContent(type, content, position, align, size, font)
Creates a Paragraph or Page in the document

**Kind**: static method of [<code>toDoc</code>](#toDoc)
**Access**: public

| Param | Type | Description |
| --- | --- | --- |
| type | <code>string</code> | Defines whether the content is a Paragraph or a Page <br/> Accepts : "paragraph" : Inserts a paragraph in the next line </br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; "page" : Inserts a paragraph in the next page <br/> <br/> Required |
| content | <code>string</code> | The content that will be created in the document <br/> Accepts : Text &#124; Stringified HTML <br/> Required |
| position | <code>number</code> | Specifies the content's position in the document <br/> Paragraphs and pages will be sorted by positon during document generation <br/> Accepts : Numbers > 0 <br/> Default position : 0 <br/> Required for pagagraphs and pages <br/> Optional if passing a whole HTML document |
| align | <code>string</code> | Specifies the content's alignemnt <br/> Accepts : "left" &#124; "center" &#124; "right" &#124; "justify" <br/> Default alignment : left <br/> Optional |
| size | <code>string</code> | Specifies the content's font size <br/> Accepts : CSS font-size <br/> Default size : 12pt/16px/1em  <br/> Optional |
| font | <code>string</code> | Specifies the content's font family <br/> Accepts : CSS font-family <br/> Default font : Times New Roman <br/> Optional |

<a name="toDoc.doc.createImage"></a>

### toDoc.doc.createImage(sectionType, imageURL, position, align, imageWidth, imageHeight)
Creates an image in the Header, Footer or Body  of the document

**Kind**: static method of [<code>toDoc</code>](#toDoc)
**Access**: public

| Param | Type | Description |
| --- | --- | --- |
| sectionType | <code>string</code> | Defines the section where the image will be created <br/> Accepts : "header" &#124; "footer" &#124; "body" <br/> Required |
| imageURL | <code>string</code> | Specifies the image's URL <br/> Accepts : image URL &#124; base64 data URL <br/> Required |
| position | <code>number</code> | Specifies the image's position in the document <br/> Images will be sorted by position along with Header/Footer/Paragraphs/Pages during document generation <br/> Accepts : Numbers > 0 <br/> Default position : 0 <br/> Optional |
| align | <code>string</code> | Defines the image's alignemnt <br/> Accepts : "left" &#124; "center" &#124; "right" <br/> Default alignment : left <br/> Optional |
| imageWidth | <code>number</code> | Specify a custom width for images <br/> Use only for resizing images <br/> Accepts : Numbers > 0 <br/> Optional |
| imageHeight | <code>number</code> | Specify a custom height for images <br/> Use only for resizing images <br/> Accepts : Numbers > 0 <br/> Optional |

<a name="toDoc.doc.createPagenumber"></a>

### toDoc.doc.createPagenumber(sectionType, format, align, size, font)
Inserts a page number in specified section of the document

**Kind**: static method of [<code>toDoc</code>](#toDoc)
**Access**: public

| Param | Type | Description |
| --- | --- | --- |
| sectionType | <code>string</code> | Defines the section where the page number is inserted in the document <br/> Accepts : "header" &#124; "footer" <br/> Required |
| format | <code>number</code> | Specifies the page number's format <br/> Accepts : 1 : Only page number </br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 2 : Page X of Y <br/> Default value : 1 <br/> Optional |
| align | <code>string</code> | Specifies the page number's alignemnt <br/> Accepts : "left" &#124; "center" &#124; "right" <br/> Default alignment : left <br/> Optional |
| size | <code>string</code> | Specifies the page number's font size <br/> Accepts : CSS font-size <br/> Default size : 12pt/16px/1em  <br/> Optional |
| font | <code>string</code> | Specifies the page number's font family <br/> Accepts : CSS font-family <br/> Default font : Times New Roman <br/> Optional |

<a name="toDoc.doc.createDocument"></a>

### toDoc.doc.createDocument(fileName, params)
Generates and saves a Word document

**Kind**: static method of [<code>toDoc</code>](#toDoc)
**Access**: public

| Param | Type | Description |
| --- | --- | --- |
| fileName | <code>string</code> | Specifies the name of the document <br/> Required |
| params | <code>object</code> | Define custom parameters for the document's layout <br/> Accepts : JSON Object <br/> Optional |

<a name="toDoc.doc.clearDocument"></a>

### toDoc.doc.clearDocument()
Clears all created document data

**Kind**: static method of [<code>toDoc</code>](#toDoc)
**Access**: public
