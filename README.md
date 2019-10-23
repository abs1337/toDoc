# toDoc

Convert HTML/Text to a Word Documnet



* * *

### toDoc.toDoc() 

Master function

**Returns**: `object`, Library object with global functions


### toDoc.doc.createHeaderFooter(headerFooter, type, content, alignment, position, nextLine) 

Creates a Header or a Footer object for the document

**Parameters**

**headerFooter**: `string`, Defines whether the object is header or footer <br/> Accepts : "header" | "footer" <br/> Required

**type**: `string`, Defines whether the object contains an image, text or HTML markup <br/> Accepts : "image" | "text" | "html" <br/> Required

**content**: `string`, Defines the object's content <br/> Accepts : stringified image URLs | stringified text | stringified HTML markup <br/> Required

**alignment**: `string`, Defines the object's alignemnt <br/> Accepts : "left" | "center"| "right" <br/> Required

**position**: `number`, Specifies the position of the object <br/> Accepts : 1++ <br/> Default value :  0 <br/> Optional

**nextLine**: `boolean`, Specifes whether the object should start in a new line <br/> Accepts : true | false <br/> Default value : false <br/> Optional



### toDoc.doc.createPage(content, number) 

Creates a Page object for the document

**Parameters**

**content**: `string`, Defines the page's content <br/> Accepts : stringified text | stringified HTML markup <br/> Required

**number**: `number`, Defines the Pages's number <br/> Accepts : 1+= <br/> Default value : 0 <br/> Required for individual pages <br/> Required if creating individual pages <br/> Optional if passing a whole HTML document



### toDoc.doc.createParagraph(content, number) 

Creates a Page object for the document

**Parameters**

**content**: `string`, Defines the page's content <br/> Accepts : stringified text | stringified HTML markup <br/> Required

**number**: `number`, Defines the Pages's number <br/> Accepts : 1+= <br/> Default value : 0 <br/> Required for individual pages <br/> Required if creating individual pages <br/> Optional if passing a whole HTML document



### toDoc.doc.createDocument(filename, params) 

Generates and saves a Word document.

**Parameters**

**filename**: `string`, Specifies the name of the document <br/> Required

**params**: `object`, Define custom parameters for the document's layout <br/> Optional




* * *










