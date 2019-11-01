/**
 * Convert HTML/Text to a Word Documnet
 * @namespace toDoc
 */
(function(window) {
    // "use strict";

    /** 
     * Master function
     * @public
     * @returns {object} Library object with global functions
     * @memberof toDoc
     */
    function toDoc() {
        var doc = {};

        /**
         * Store default document settings
         * @private
         * @memberof toDoc
         */
        var oSettings = {
            pageSizeX: "8.5in",
            pageSizeY: "11in",
            marginTop: "1in",
            marginBottom: "1in",
            marginLeft: "1in",
            marginRight: "1in",
            headerMargin: "0.5in",
            footerMargin: "0.5in",
        };

        /** 
         * Store Paragraph, Page and Section data
         * @private
         * @memberof toDoc
         */
        var oData = {
            // Header
            aHeader: [],
            // Footer
            aFooter: [],
            // Page / Paragraph
            aContent: [],

            // HTML tags and markup
            break: "<br/>",
            pageBreak: "<br clear='all' style='page-break-before:always' />",
            paraStart: "<p>",
            paraEnd: "</p>",
            fontStart: "<font size='1'>",
            fontEnd: "</font>",
            pageNum_1: " <span style='mso-field-code: PAGE '></span>",
            pageNum_2: " Page <span style='mso-field-code: PAGE '></span> of <span style='mso-field-code: NUMPAGES '></span>",
        };

        /** 
         * Creates a Header or a Footer Section in the document
         * @public
         * @param {string} sectionType - Specifies whether the content is for the header or footer of the document <br/> Accepts : "header" &#124; "footer" <br/> Required
         * @param {string} contentType - Defines whether the section's content is Text or stringified HTML markup <br/> Accepts : "text" &#124; "html" <br/> Required
         * @param {string} content -  The content that will be created in section <br/> Accepts : stringified text &#124; stringified HTML markup <br/> Required
         * @param {number} position - Specifies the content's position in the section <br/> Accepts : 1++ <br/> Default position: 0 <br/> Optional
         * @param {string} align - Defines the section content's alignemnt <br/> Accepts : "left" &#124; "center" &#124; "right" &#124; "justify" <br/> Default alignment: left <br/> Optional
         * @param {string} size - Specifies the content's font size <br/> Accepts : CSS font-size <br/> Default size: 12pt/16px/1em  <br/> Optional
         * @param {string} font - Specifies the content's font family <br/> Accepts : CSS font-family <br/> Default font: Times New Roman <br/> Optional
         * @memberof toDoc
         */
        doc.createSection = function(sectionType, contentType, content, position, align, size, font) {

            // Store section data and settings as an object
            var sectionObj = {
                "sContentType": "",
                "sContent": "",
                "sPosition": 0,
                "sAlign": "",
                "sSize": "",
                "sFont": "",
                "iAlign": "",
                "iWidth": "",
                "iHeight": "",
            };

            // Validate and set position
            if (!position) {
                sectionObj.sPosition = 0;
            } else if (position && Number.isInteger(position) && position > 0) {
                sectionObj.sPosition = position;
            } else {
                console.error(position + " is an invalid header position, expected number");
                return;
            }

            // Validate and set alignment
            if (!align) {
                sectionObj.sAlign = "left";
            } else if (align && (align === "left" || align === "center" || align === "right" || align === "justify")) {
                sectionObj.sAlign = align;
            } else {
                console.error(align + " is an invalid alignment for image, expected 'left', 'top', 'middle', 'left' or 'right'");
                return;
            }

            // Validate and set font size
            var fontSizeRegex = /(?:[^\d]\.| |^)((?:\d+\.)?\d+) *pt$|px$|em$/; // Regex for number followed by 'pt', 'px' or 'em'
            if (!size) {
                sectionObj.sSize = "16px";
            } else if (size && typeof(size) == "string" && fontSizeRegex.test(size)) {
                sectionObj.sSize = size;
            } else {
                console.error(size + " is an invalid header position, expected number");
                return;
            }

            // Validate and set font family
            if (!font) {
                sectionObj.sFont = "";
            } else if (font && typeof(font) == "string") {
                sectionObj.sFont = font;
            } else {
                console.error(font + " is an invalid header position, expected number");
                return;
            }

            // Crete section content based on type
            switch (contentType) {
                // Create Header Text
                case "text":
                    if (content.length > 0 && typeof(content) == "string") {
                        sectionObj.sContent = content;
                        sectionObj.sContentType = contentType;
                    } else {
                        console.error(content + "is not an valid value for the parameter 'content', expected string");
                        return;
                    }
                    break;
                    // Create Header HTML Markup
                case "html":
                    var htmlRegex = /<\/?[a-z][\s\S]*>/i;
                    // Check for valid HTML
                    if (content.length > 0 && htmlRegex.test(content)) {
                        sectionObj.sContent = content;
                        sectionObj.sContentType = contentType;
                    } else {
                        console.error("Invalid HTML markup");
                        return;
                    }
                    break;
                    // Error
                default:
                    console.error(contentType + " is an invalid content type, expected 'string' or 'html'");
                    return;
            }

            // Create section based on type - Header / Footer and push to respestive array
            if (sectionType.length > 0 && typeof(sectionType) == "string" && (sectionType === "header" || sectionType === "footer")) {
                if (sectionType === "header") {
                    oData.aHeader.push(sectionObj);
                } else if (sectionType === "footer") {
                    oData.aFooter.push(sectionObj);
                }
            } else {
                console.error(sectionType + " is an invalid section type, expected 'header' or 'footer'");
                return;
            }
        };

        /** 
         * Stitch sections during document generation
         * @private
         * @returns {string} Returns stringified markup of header and footer content
         * @memberof toDoc
         */
        var getSection = function(sectionArray) {
            var sections = sectionArray;
            var sectionString = "";
            // Sort by ascending position
            if (sections.length > 1) {
                sections.sort(function(a, b) {
                    var posA = a["cPosition"];
                    var posB = b["cPosition"];
                    return (posA < posB) ? -1 : (posA > posB) ? 1 : 0;
                });
            }
            // Stich section elements together
            if (sections.length > 0) {
                sections.forEach(function(i) {
                    if (i.sContentType === "text") {
                        sectionString = sectionString + "<p align='" + i.sAlign + "' style='font-size : " + i.sSize + "; font-family : " + i.sFont + ", Times New Roman;'>" + i.sContent + oData.paraEnd;
                    } else if (i.sContentType === "image" || i.sContentType === "html") {
                        sectionString = sectionString + i.sContent;
                    }
                });
            }
            return sectionString;
        };

        /** 
         * Creates a Paragraph or Page in the document
         * @public
         * @param {string} type - Specify whether the content is a Paragraph or a new Page <br/> Accepts : "paragraph" &#124; "page" <br/> Required
         * @param {string} content - The content that will be created as a Paragraph or Page <br/> Accepts : stringified text &#124; stringified HTML markup <br/> Required
         * @param {number} position - Defines the content's position in the document <br/> Accepts : 1++ <br/> Default value : 0 <br/> Required for pagagraphs and pages <br/> Optional if passing a whole HTML document
         * @param {string} align - Defines the content's alignemnt <br/> Accepts : "left" &#124; "center" &#124; "right" &#124; "justify" <br/> Default alignment: left <br/> Optional
         * @param {string} string - Specifies the content's font size <br/> Accepts : CSS font-size <br/> Default size: 12pt/16px/1em  <br/> Optional
         * @param {string} font - Specifies the content's font family <br/> Accepts : CSS font-family <br/> Default font: Times New Roman <br/> Optional
         * @memberof toDoc
         */
        doc.createContent = function(type, content, position, align, size, font) {

            var contentObj = {
                "cContent": "",
                "cPosition": 0,
                "cType": "",
                "cAlign": "",
                "cSize": "",
                "cFont": "",
                "iAlign": "",
                "iWidth": "",
                "iHeight": "",
            };

            // Validate and set position
            if (!position) {
                contentObj.sPosition = 0;
            } else if (position && Number.isInteger(position) && position > 0) {
                // Check for duplicate content position
                if (oData.aContent.length > 0) {
                    oData.aContent.every(function(i) {
                        if (i.hasOwnProperty("cPosition")) {
                            if (i.cPosition == position) {
                                console.error("Content with position: " + i.cPosition + " has already been defined, enter new position for " + content);
                                return false;
                            } else {
                                contentObj.sPosition = position;
                                return true;
                            }
                        }
                    });
                }
            } else {
                console.error(position + " is an invalid header position, expected number");
                return;
            }

            // Validate and set alignment
            if (!align) {
                contentObj.cAlign = "left";
            } else if (align && (align === "left" || align === "center" || align === "right" || align === "justify")) {
                contentObj.cAlign = align;
            } else {
                console.error(align + " is an invalid alignment for image, expected 'left', 'top', 'middle', 'left' or 'right'");
                return;
            }

            // Validate and set font size
            var fontSizeRegex = /(?:[^\d]\.| |^)((?:\d+\.)?\d+) *pt$|px$|em$/; // Regex for number followed by 'pt', 'px' or 'em'
            if (!size) {
                contentObj.cSize = "16px";
            } else if (size && typeof(size) == "string" && fontSizeRegex.test(size)) {
                contentObj.cSize = size;
            } else {
                console.error(size + " is an invalid header position, expected number");
                return;
            }

            // Validate and set font face
            if (!font) {
                contentObj.cFont = "";
            } else if (font && typeof(font) == "string") {
                contentObj.cFont = font;
            } else {
                console.error(font + " is an invalid header position, expected number");
                return;
            }

            switch (type) {

                // Create a Paragraph with text or HTML markup
                case "paragraph":
                    if (typeof(content) == "string" && content.length > 0) {
                        contentObj.cContent = content;
                        // Check if content is Text or HTML markup
                        var htmlRegex = /<\/?[a-z][\s\S]*>/i;
                        if (htmlRegex.test(content)) {
                            contentObj.cType = "html";
                        } else {
                            contentObj.cType = type;
                        }
                    } else {
                        console.error(content + " is not an valid value for the parameter 'content', expected string");
                        return;
                    }
                    break;

                    // Create a page with HTML markup
                case "page":
                    if (typeof(content) == "string" && content.length > 0) {
                        contentObj.cContent = content;
                        // Check for HTML markup
                        var htmlRegex = /<\/?[a-z][\s\S]*>/i;
                        if (htmlRegex.test(content)) {
                            contentObj.cType = "html_page";
                        } else {
                            contentObj.cType = type;
                        }
                    } else {
                        console.error(content + "is not an valid value for the parameter 'content', expected string");
                        return;
                    }
                    break;

                    // Error
                default:
                    console.error(type + " is an invalid content type, expected 'image', 'string' or 'html'");
                    return;
            }

            // Push to array with other contents
            oData.aContent.push(contentObj);
        };

        /** 
         * Stitch Page during document generation
         * @private
         * @returns {string} Returns stringified markup of all pages
         * @memberof toDoc
         */
        var getContents = function() {
            var content = oData.aContent;

            // Sort by ascending position
            if (content.length > 1) {
                content.sort(function(a, b) {
                    var posA = a["cPosition"];
                    var posB = b["cPosition"];
                    return (posA < posB) ? -1 : (posA > posB) ? 1 : 0;
                });
            }

            // Stitch content together
            if (content.length > 0) {
                var contentString = "";
                content.forEach(function(i, index) {

                    // Stitch markup based on content type
                    if (i.cType === "paragraph") {

                        contentString = contentString + "<p align='" + i.cAlign + "' style='font-size : " + i.cSize + "; font-family : " + i.cFont + ", Times New Roman;'>" + i.cContent + oData.paraEnd;

                    } else if (i.cType === "page") {

                        contentString = oData.pageBreak + contentString + i.cContent;

                    } else if (i.cType === "image" || i.cType === "html") {

                        contentString = contentString + oData.break+i.cContent + oData.break;

                    } else if (i.cType === "html_page") {

                        contentString = oData.pageBreak + contentString + i.cContent;

                    }

                    // Create content in next line
                    if (i.contentNextLine) {
                        contentString = contentString + oData.break;
                    }

                });
                return contentString;
            } else {
                return "";
            }
        };

        /** 
         * Creates an image in the Header, Footer or Body  of the document
         * @public
         * @param {string} sectionType - Specifies where the image will be inserted in the document <br/> Accepts : "header" &#124; "footer" &#124; "body" <br/> Required
         * @param {string} imageURL - Specifies the image's URL <br/> Accepts : stringified image URL <br/> Required
         * @param {number} position - Specifies the image's position in the document <br/> Accepts : 1++ <br/> Default value: 0 <br/> Optional
         * @param {string} align - Defines the image's alignemnt <br/> Accepts : "left" &#124; "center" &#124; "right" <br/> Default alignment: left <br/> Optional
         * @param {number} imageWidth - Specify a custom width for images <br/> Only works for type: "image" <br/> Optional
         * @param {number} imageHeight - Specify a custom height for images <br/> Only works for type: "image" <br/> Optional
         * @memberof toDoc
         */
        doc.createImage = function(sectionType, imageURL, position, align, imageWidth, imageHeight) {

            // Store section data and settings
            var sectionObj = {
                "sContentType": "",
                "sContent": "",
                "sPosition": 0,
                "sAlign": "",
                "sSize": "",
                "sFont": "",
                "iAlign": "",
                "iWidth": "",
                "iHeight": "",
            };

            var contentObj = {
                "cContent": "",
                "cPosition": 0,
                "cType": "",
                "cAlign": "",
                "cSize": "",
                "cFont": "",
                "iAlign": "",
                "iWidth": "",
                "iHeight": "",
            };

            var imageObj = {};

            // Create image for header/footer or body
            if (sectionType.length > 0 && typeof(sectionType) == "string" && (sectionType === "header" || sectionType === "footer" || sectionType === "body")) {
                if (sectionType === "header" || sectionType === "footer") {
                    imageObj = sectionObj;
                } else if (sectionType === "body") {
                    imageObj = contentObj;
                }
            } else {
                console.error(sectionType + " is an invalid section type, expected 'header' or 'footer'");
                return;
            }

            // Validate and set position
            if (!position) {
                imageObj.sPosition = 0;
            } else if (position && Number.isInteger(position) && position > 0) {
                imageObj.sPosition = position;
            } else {
                console.error(position + " is an invalid header position, expected number");
                return;
            }

            // Validate and set alignment
            if (!align) {
                imageObj.sAlign = "left";
                imageObj.iAlign = "left";
            } else if (align && (align === "left" || align === "center" || align === "right" || align === "justify")) {
                imageObj.sAlign = align;
                imageObj.iAlign = align;
            } else {
                console.error(align + " is an invalid alignment for image, expected 'left', 'top', 'middle', 'left' or 'right'");
                return;
            }

            // Validate and set image width
            if (!imageWidth) {
                imageObj.iWidth = "";
            } else if (imageWidth && Number.isInteger(imageWidth) && imageWidth > 0) {
                imageObj.iWidth = imageWidth;
            } else {
                console.error(imageWidth + " is an invalid header position, expected number");
                return;
            }

            // Validate and set image height
            if (!imageHeight) {
                imageObj.iHeight = "";
            } else if (imageHeight && Number.isInteger(imageHeight) && imageHeight > 0) {
                imageObj.iHeight = imageHeight;
            } else {
                console.error(imageHeight + " is an invalid header position, expected number");
                return;
            }

            // Validate URL and create image
            var urlRegex = /^(?:(?:https?|ftp):\/\/)?(?:(?!(?:10|127)(?:\.\d{1,3}){3})(?!(?:169\.254|192\.168)(?:\.\d{1,3}){2})(?!172\.(?:1[6-9]|2\d|3[0-1])(?:\.\d{1,3}){2})(?:[1-9]\d?|1\d\d|2[01]\d|22[0-3])(?:\.(?:1?\d{1,2}|2[0-4]\d|25[0-5])){2}(?:\.(?:[1-9]\d?|1\d\d|2[0-4]\d|25[0-4]))|(?:(?:[a-z\u00a1-\uffff0-9]-*)*[a-z\u00a1-\uffff0-9]+)(?:\.(?:[a-z\u00a1-\uffff0-9]-*)*[a-z\u00a1-\uffff0-9]+)*(?:\.(?:[a-z\u00a1-\uffff]{2,})))(?::\d{2,5})?(?:\/\S*)?$/;
            // Check for valid URL
            if (urlRegex.test(imageURL)) {
                var imgUrl = imageURL,
                    imgAlign = imageObj.iAlign,
                    imgWidth = imageObj.iHeight,
                    imgHeight = imageObj.iWidth;
                // Get image as DATA URL
                getImage(imgUrl, imgAlign, imgWidth, imgHeight, function(imgMarkup) {
                    imageObj.sContent = imgMarkup;
                    imageObj.cType = "image";
                    imageObj.sContentType = "image";
                });
            } else {
                console.error(imageURL + "is not an valid value for the parameter 'imageURL', expected image url/path");
                return;
            }

            // Create section based on type - Header / Footer / Body and push to respestive array
            if (sectionType === "header") {
                oData.aHeader.push(imageObj);
            } else if (sectionType === "footer") {
                oData.aFooter.push(imageObj);
            } else if (sectionType === "body") {
                oData.aContent.push(imageObj);
            }
        };

        /** 
         * Creates and stores an image as a local data URL from image's URL
         * @private
         * @param {string} url - URL or path of the image
         * @param {string} imgAlign - Specifies the image's alignment
         * @param {string} imgWidth - Defines the image's width
         * @param {string} imgHeight - Defines the image's height
         * @param {function} callBack - Callback function return data URL after image loads
         * @returns {string} Returns a stringified HTML markup of the image
         * @memberof toDoc
         */
        var getImage = function(url, imgAlign, imgWidth, imgHeight, callBack) {
            var image = new Image();
            image.crossOrigin = "anonymous"

            // Execute after image loads
            image.onload = function() {

                var nHeight = image.height;
                var nWidth = image.width;

                // Set image to canvas to get data URL
                if (image != null && (nWidth && nHeight)) {

                    var imgCanvas = document.createElement("canvas"),
                        imgContext = imgCanvas.getContext("2d");

                    // Make sure canvas is as big as the picture
                    imgCanvas.width = image.width;
                    imgCanvas.height = image.height;

                    // Draw image into canvas element
                    imgContext.drawImage(image, 0, 0, image.width, image.height);

                    // Get canvas contents as a data URL
                    var imgAsDataURL = imgCanvas.toDataURL("image/png");

                    var imgHTML = "<div align='" + imgAlign + "'><img width='" + imgWidth + "' height='" + imgHeight + "' src='" + imgAsDataURL + "' crossOrigin='anonymous'></img>"

                    // Return data URL in HTML tags
                    if (callBack) {
                        callBack(imgHTML);
                    } else {
                        return imgHTML;
                    }
                } else {
                    console.error("cannot reach image url");
                    return "";
                }
            }
            
            // Load image
            image.src = url;

            // Try Handling CORS error.
            /*var request = new XMLHttpRequest();
            request.open("GET", url);
            request.onload = request.onerror = function() {
                for (var responseText = request.responseText, responseTextLen = responseText.length, binary = "", i = 0; i < responseTextLen; ++i) {
                  binary += String.fromCharCode(responseText.charCodeAt(i) & 255)
                }
                var src = 'data:image/jpeg;base64,'+ window.btoa(binary);

                var imgHTML = "<div align='" + imgAlign + "'><img width='" + imgWidth + "' height='" + imgHeight + "' src='" + src + "' crossOrigin='anonymous'></img>"

                // Return data URL in HTML tags
                if (callBack) {
                    callBack(imgHTML);
                } else {
                    return imgHTML;
                }
            };
            request.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
            request.send("");*/
        };

        /** 
         * Inserts page number in specified section of the document
         * @public
         * @param {string} sectionType - Specifies where the page number is inserted in the document <br/> Accepts : "header" &#124; "footer" <br/> Required
         * @param {number} format - Specifies the page number's format <br/> Accepts : 1 &#124; 2 <br/> Default value: 1 <br/> Optional
         * @param {string} align - Defines the page number's alignemnt <br/> Accepts : "left" &#124; "center" &#124; "right" <br/> Default alignment: left <br/> Optional
         * @param {string} string - Specifies the page number's font size <br/> Accepts : CSS font-size <br/> Default size: 12pt/16px/1em  <br/> Optional
         * @param {string} font - Specifies the page number's font family <br/> Accepts : CSS font-family <br/> Default font: Times New Roman <br/> Optional
         * @memberof toDoc
         */
        doc.createPagenum = function(sectionType, format, align, size, font) {

            // Store section data and settings
            var sectionObj = {
                "sContentType": "",
                "sContent": "",
                "sPosition": 0,
                "sAlign": "",
                "sSize": "",
                "sFont": "",
                "iAlign": "",
                "iWidth": "",
                "iHeight": "",
            };

            // Validate and set alignment
            if (!align) {
                sectionObj.sAlign = "left";
            } else if (align && (align === "left" || align === "center" || align === "right")) {
                sectionObj.sAlign = align;
            } else {
                console.error(align + " is an invalid alignment for image, expected 'left', 'center' or 'right'");
                return;
            }

            // Validate text's font size
            var fontSizeRegex = /(?:[^\d]\.| |^)((?:\d+\.)?\d+) *pt$|px$|em$/; // Regex for number followed by 'pt', 'px' or 'em'
            if (!size) {
                sectionObj.sSize = "16px";
            } else if (size && typeof(size) == "string" && fontSizeRegex.test(size)) {
                sectionObj.sSize = size;
            } else {
                console.error(size + " is an invalid header position, expected number");
                return;
            }

            // Validate text's font face
            if (!font) {
                sectionObj.sFont = "";
            } else if (font && typeof(font) == "string") {
                sectionObj.sFont = font;
            } else {
                console.error(font + " is an invalid font style");
                return;
            }

            // Format page number
            if (!format || format === 1) {
                sectionObj.sContent = "<p align='" + sectionObj.sAlign + "' style='font-size : '" + sectionObj.sSize + "; font-family : " + sectionObj.sFont + ", Times New Roman;'>" + oData.pageNum_1 + oData.paraEnd;
            } else if (format === 2) {
                sectionObj.sContent = "<p align='" + sectionObj.sAlign + "' style='font-size : '" + sectionObj.sSize + "; font-family : " + sectionObj.sFont + ", Times New Roman;'>" + oData.pageNum_2 + oData.paraEnd;
            } else {
                console.error(format + " is an invalid format for page numbers, expected number 1/2");
                return; 
            }

            // Insert page number in specified section and push to respestive array
            if (sectionType.length > 0 && typeof(sectionType) == "string" && (sectionType === "header" || sectionType === "footer")) {
                if (sectionType === "header") {
                    oData.aHeader.push(sectionObj);
                } else if (sectionType === "footer") {
                    oData.aFooter.push(sectionObj);
                }
            } else {
                console.error(sectionType + " is an invalid section type, expected 'header' or 'footer'");
                return;
            }

        };


        /** 
         * Clears all section and content data
         * @public
         * @param {string} fileName - Specifies the name of the document <br/> Required
         * @param {object} params - Define custom parameters for the document's layout <br/> Optional
         * @memberof toDoc
         */
        doc.clearDocument = function() {
            // Header
            oData.aHeader = [];
            // Footer
            oData.aFooter = [];
            // Page / Paragraph
            oData.aContent = [];
        };

        /** 
         * Generates and saves a Word document.
         * @public
         * @param {string} fileName - Specifies the name of the document <br/> Required
         * @param {object} params - Define custom parameters for the document's layout <br/> Optional
         * @memberof toDoc
         */
        doc.createDocument = function(fileName, params) {
            var docParams = {};

            // Validate object parameters
            if (params && !params instanceof Array) {

                Object.keys(params).forEach(function(i) {
                    // Check for number followed by 'cm' or 'in'
                    var unitRegex = /(?:[^\d]\.| |^)((?:\d+\.)?\d+) *cm$|in$/;
                    if (!unitRegex.test(params[i])) {
                        console.error(params[i] + " in an invalid value for parameter: " + i);
                        return;
                    }
                });
                docParams = params;
            } else {
                docParams = oSettings;
            }

            // HTML markup for Word Document
            var documentMarkup = "<html xmlns:o='urn:schemas-microsoft-com:office:office'"+
                "xmlns:w='urn:schemas-microsoft-com:office:word'"+
                "xmlns='http://www.w3.org/TR/REC-html40'>"+
                "<head>"+
                    "<style>"+
                        "p.MsoHeader, li.MsoHeader, div.MsoHeader{"+
                            "margin:0in;"+
                            "margin-top:.0001pt;"+
                            "mso-pagination:widow-orphan;"+
                            "tab-stops:center 3.0in right 6.0in;"+
                        "}"+
                        "p.MsoFooter, li.MsoFooter, div.MsoFooter{"+
                            "margin:0in;"+
                            "margin-bottom:.0001pt;"+
                            "mso-pagination:widow-orphan;"+
                            "tab-stops:center 3.0in right 6.0in;"+
                            "font-size:18.0pt;" +
                        "}"+
                        "@page Section1{"+
                            "size:" + docParams.pageSizeX + " " + docParams.pageSizeY + ";"+
                            "margin:" + docParams.marginTop + " " + docParams.marginBottom + " " + docParams.marginLeft + " " + docParams.marginRight + ";"+
                            "mso-header-margin:" + docParams.headerMargin + ";"+
                            "mso-header:h1;"+
                            "mso-footer:f1;"+
                            "mso-footer-margin:" + docParams.footerMargin + ";"+
                            "mso-paper-source:0;"+
                        "}"+
                        "div.Section1{"+
                            "page:Section1;"+
                        "}"+
                        //Style to stop Header and Footer repeating
                        "table#hrdftrtbl{"+
                            "margin:0in 0in 0in 900in;"+
                        "}"+
                    "</style>"+
                "</head>"+
                "<body>"+
                    //Content
                    "<div class='Section1'>"+ // Section1
                        //Pages 
                        getContents() +
                        //Header and Footer
                        "<table id='hrdftrtbl' border='1' cellspacing='0' cellpadding='0'>"+
                            "<tr>"+
                                "<td>"+
                                    //Header
                                    "<div style='mso-element:header' id='h1' >"+
                                        "<p class='MsoHeader'>"+
                                            "<table border='0' width='100%'>"+
                                                //"<tr>"+
                                                    // Header Left Align
                                                    //"<td width='30%'>"+
                                                        getSection(oData.aHeader) +
                                                    //"</td>"+
                                                    // Header Center Align
                                                    //"<td align='center' width='40%'>"+
                                                    //    getSection(oData.aHeaderCenter) +
                                                    //"</td>"+
                                                    // Header Right Align
                                                    //"<td align='right' width='30%'>"+
                                                    //    getSection(oData.aHeaderRight) +
                                                    //"</td>"+
                                                //"</tr>"+
                                            "</table>"+
                                        "</p>"+
                                    "</div>"+
                                "</td>"+
                                "<td>"+
                                    //Footer
                                    "<div style='mso-element:footer' id='f1'>"+
                                        "<p class='MsoFooter'>"+
                                            "<table width='100%' border='0' cellspacing='0' cellpadding='0'>"+
                                                //"<tr>"+
                                                    // Footer Left Align
                                                    //"<td width='30%'>"+
                                                        getSection(oData.aFooter) +
                                                    //"</td>"+
                                                    // Footer Center Align
                                                    //"<td align='center' width='40%'>"+
                                                    //    getSection(oData.aFooterCenter) +
                                                    //"</td>"+
                                                    // Footer Right Align
                                                    //"<td align='right' width='30%'>"+
                                                    //    getSection(oData.aFooterRight) +
                                                    //    "Page <span style='mso-field-code: PAGE '></span> of <span style='mso-field-code: NUMPAGES '></span>"+
                                                    //"</td>"+
                                                //"</tr>"+
                                            "</table>"+
                                        "</p>"+
                                    "</div>"+
                                "</td>"+
                            "</tr>"+
                        "</table>"+
                    "</div>"+
                    // End Section 1
                "</body>"+
            "</html>"

            var blob = new Blob(['\ufeff', documentMarkup], {
                type: 'application/msword'
            });

            // Get link URL
            var url = URL.createObjectURL(blob);

            // Set file name
            if (!fileName) {
                fileName = fileName ? fileName + '.doc' : 'document.doc';
            }

            // Create download link element
            var downloadLink = document.createElement("a");
            document.body.appendChild(downloadLink);

            // Check for IE 10/11
            if (navigator.msSaveOrOpenBlob) {
                navigator.msSaveOrOpenBlob(blob, fileName);
            } else {
                // Create a link to the file
                downloadLink.href = url;
                // Set the file name
                downloadLink.download = fileName;
                // Trigger the function
                downloadLink.click();
            }
            document.body.removeChild(downloadLink);

            // Clear previous document's data
            doc.clearDocument();
        };

        return doc;
    }

    // Set Lib globally and save to window
    if (typeof(window.toDoc) === 'undefined') {
        window.toDoc = toDoc();
    }

})(window);