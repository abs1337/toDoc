## toDoc : <code>object</code>
Convert HTML/Text to a Word Documnet

**Kind**: global namespace

* [toDoc](#toDoc) : <code>object</code>
    * [.toDoc()](#toDoc.toDoc) ⇒ <code>object</code>
    * [.doc.createHeaderFooter(headerFooter, type, content, alignment, position, nextLine)](#toDoc.doc.createHeaderFooter)
    * [.doc.createPage(content, number)](#toDoc.doc.createPage)
    * [.doc.createParagraph(content, number)](#toDoc.doc.createParagraph)
    * [.doc.createDocument(filename, params)](#toDoc.doc.createDocument)

<a name="toDoc.toDoc"></a>

### toDoc.toDoc() ⇒ <code>object</code>
Master function

**Kind**: static method of [<code>toDoc</code>](#toDoc)
**Returns**: <code>object</code> - Library object with global functions
**Access**: public
<a name="toDoc.doc.createHeaderFooter"></a>

### toDoc.doc.createHeaderFooter(headerFooter, type, content, alignment, position, nextLine)
Creates a Header or a Footer object for the document

**Kind**: static method of [<code>toDoc</code>](#toDoc)
**Access**: public

| Param | Type | Description |
| --- | --- | --- |
| headerFooter | <code>string</code> | Defines whether the object is header or footer <br/> Accepts : "header" | "footer" <br/> Required |
| type | <code>string</code> | Defines whether the object contains an image, text or HTML markup <br/> Accepts : "image" | "text" | "html" <br/> Required |
| content | <code>string</code> | Defines the object's content <br/> Accepts : stringified image URLs | stringified text | stringified HTML markup <br/> Required |
| alignment | <code>string</code> | Defines the object's alignemnt <br/> Accepts : "left" | "center"| "right" <br/> Required |
| position | <code>number</code> | Specifies the position of the object <br/> Accepts : 1++ <br/> Default value :  0 <br/> Optional |
| nextLine | <code>boolean</code> | Specifes whether the object should start in a new line <br/> Accepts : true | false <br/> Default value : false <br/> Optional |

<a name="toDoc.doc.createPage"></a>

### toDoc.doc.createPage(content, number)
Creates a Page object for the document

**Kind**: static method of [<code>toDoc</code>](#toDoc)
**Access**: public

| Param | Type | Description |
| --- | --- | --- |
| content | <code>string</code> | Defines the page's content <br/> Accepts : stringified text | stringified HTML markup <br/> Required |
| number | <code>number</code> | Defines the Pages's number <br/> Accepts : 1+= <br/> Default value : 0 <br/> Required for individual pages <br/> Required if creating individual pages <br/> Optional if passing a whole HTML document |

<a name="toDoc.doc.createParagraph"></a>

### toDoc.doc.createParagraph(content, number)
Creates a Page object for the document

**Kind**: static method of [<code>toDoc</code>](#toDoc)
**Access**: public

| Param | Type | Description |
| --- | --- | --- |
| content | <code>string</code> | Defines the page's content <br/> Accepts : stringified text | stringified HTML markup <br/> Required |
| number | <code>number</code> | Defines the Pages's number <br/> Accepts : 1+= <br/> Default value : 0 <br/> Required for individual pages <br/> Required if creating individual pages <br/> Optional if passing a whole HTML document |

<a name="toDoc.doc.createDocument"></a>

### toDoc.doc.createDocument(filename, params)
Generates and saves a Word document.

**Kind**: static method of [<code>toDoc</code>](#toDoc)
**Access**: public

| Param | Type | Description |
| --- | --- | --- |
| filename | <code>string</code> | Specifies the name of the document <br/> Required |
| params | <code>object</code> | Define custom parameters for the document's layout <br/> Optional |
