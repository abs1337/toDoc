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
            pageNumber1: "<p align='center' style='margin : 0;''><font size='1'> Page <span style='mso-field-code: PAGE '></span> of <span style='mso-field-code: NUMPAGES '></span> </font></p>"
        };

        /** 
         * Creates a Header or Footer Section in the document
         * @public
         * @param {string} sectionType - Defines whether the section is a header or footer <br/> Accepts : "header", "footer" <br/> Required
         * @param {string} contentType - Defines whether the section text or stringified HTML markup <br/> Accepts : "text", "html" <br/> Required
         * @param {string} content - Defines the sections's content <br/> Accepts : stringified text, stringified HTML markup <br/> Required
         * @param {string} align - Defines the section content's alignemnt <br/> Accepts : "left", "center", "right", "justify" <br/> Default alignment: left <br/> Optional
         * @param {number} position - Specifies the content's position in the section <br/> Accepts : 1++ <br/> Default position: 0 <br/> Optional
         * @param {number} size - Specify <br/> Default size: "" <br/> Optional
         * @param {string} font - Specify <br/> Default font: Times New Roman <br/> Optional
         * @memberof toDoc
         */
        doc.createSection = function(sectionType, contentType, content, align, position, size, font) {

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
                sectionObj.iAlign = "left";
            } else if (align && (align === "left" || align === "center" || align === "right" || align === "justify")) {
                sectionObj.sAlign = align;
                sectionObj.iAlign = align;
            } else {
                console.error(align + " is an invalid alignment for image, expected 'left', 'top', 'middle', 'left' or 'right'");
                return;
            }

            // Validate and set position
            if (!position) {
                sectionObj.sPosition = 0;
            } else if (position && Number.isInteger(position) && position > 0) {
                sectionObj.sPosition = position;
            } else {
                console.error(position + " is an invalid header position, expected number");
                return;
            }

            // Validate text's font size
            if (!size) {
                sectionObj.sSize = "";
            } else if (size && Number.isInteger(size) && size > 0) {
                sectionObj.iWidth = size;
            } else {
                console.error(sSize + " is an invalid header position, expected number");
                return;
            }

            // Validate text's font face
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
         * Creates an image in the Header or Footer Section of the document
         * @public
         * @param {string} sectionType - Defines whether the section is a header or footer <br/> Accepts : "header", "footer" <br/> Required
         * @param {string} imageURL - Defines the sections's content <br/> Accepts : stringified image URLs, stringified text, stringified HTML markup <br/> Required
         * @param {string} align - Defines the section content's alignemnt <br/> Accepts : "left", "center", "right", "justify" <br/> Default alignment: left <br/> Optional
         * @param {number} position - Specifies the content's position in the section <br/> Accepts : 1++ <br/> Default value: 0 <br/> Optional
         * @param {number} imageWidth - Specify custom width for images <br/> Only works for type: "image" <br/> Optional
         * @param {number} imageHeight - Specify custom height for images <br/> Only works for type: "image" <br/> Optional
         * @memberof toDoc
         */
        doc.createSectionImage = function(sectionType, imageURL, align, position, imageWidth, imageHeight) {

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
                sectionObj.iAlign = "left";
            } else if (align && (align === "left" || align === "center" || align === "right" || align === "justify")) {
                sectionObj.sAlign = align;
                sectionObj.iAlign = align;
            } else {
                console.error(align + " is an invalid alignment for image, expected 'left', 'top', 'middle', 'left' or 'right'");
                return;
            }

            // Validate and set position
            if (!position) {
                sectionObj.sPosition = 0;
            } else if (position && Number.isInteger(position) && position > 0) {
                sectionObj.sPosition = position;
            } else {
                console.error(position + " is an invalid header position, expected number");
                return;
            }

            // Validate and set image width
            if (!imageWidth) {
                sectionObj.iWidth = "";
            } else if (imageWidth && Number.isInteger(imageWidth) && imageWidth > 0) {
                sectionObj.iWidth = imageWidth;
            } else {
                console.error(imageWidth + " is an invalid header position, expected number");
                return;
            }

            // Validate and set image height
            if (!imageHeight) {
                sectionObj.iHeight = "";
            } else if (imageHeight && Number.isInteger(imageHeight) && imageHeight > 0) {
                sectionObj.iHeight = imageHeight;
            } else {
                console.error(imageHeight + " is an invalid header position, expected number");
                return;
            }

            // Validate URL and create image
            var urlRegex = /^(?:(?:https?|ftp):\/\/)?(?:(?!(?:10|127)(?:\.\d{1,3}){3})(?!(?:169\.254|192\.168)(?:\.\d{1,3}){2})(?!172\.(?:1[6-9]|2\d|3[0-1])(?:\.\d{1,3}){2})(?:[1-9]\d?|1\d\d|2[01]\d|22[0-3])(?:\.(?:1?\d{1,2}|2[0-4]\d|25[0-5])){2}(?:\.(?:[1-9]\d?|1\d\d|2[0-4]\d|25[0-4]))|(?:(?:[a-z\u00a1-\uffff0-9]-*)*[a-z\u00a1-\uffff0-9]+)(?:\.(?:[a-z\u00a1-\uffff0-9]-*)*[a-z\u00a1-\uffff0-9]+)*(?:\.(?:[a-z\u00a1-\uffff]{2,})))(?::\d{2,5})?(?:\/\S*)?$/;
            // Check for valid URL
            if (urlRegex.test(content)) {
                var imgUrl = content,
                    imgAlign = sectionObj.iAlign,
                    imgHeight = sectionObj.iWidth,
                    imgWidth = sectionObj.iHeight;
                // Get image as DATA URL
                getImage(imgUrl, imgAlign, imgWidth, imgHeight, function(imgMarkup) {
                    sectionObj.sContent = imgMarkup;
                    sectionObj.sContentType = contentType;
                });
            } else {
                console.error(content + "is not an valid value for the parameter 'content', expected image url/path");
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
                sections.sort(function(a,b){
                    var posA = a["cPosition"];
                    var posB = b["cPosition"];
                    return (posA < posB) ? -1 : (posA > posB) ? 1 : 0;
                });
            }
            // Stich section elements together
            if (sections.length > 0) {
                sections.forEach(function(i) {
                    if (i.sContentType === "text") {
                        sectionString = sectionString + "<p align='" + i.sAlign + "'>" + "<font size='" + i.sSize + "' face ='" + i.sFont + "'>" + i.sContent + oData.fontEnd + oData.paraEnd;
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
         * @param {string} type - Specify whether the content is a Paragraph, a Page or an Image <br/> Accepts : "paragraph", "image", "page" <br/> Required
         * @param {string} content - Defines the document's content <br/> Accepts : stringified text, stringified HTML markup, stringified image URLs <br/> Required
         * @param {number} position - Defines the content's position in the document <br/> Accepts : 1++ <br/> Default value : 0 <br/> Required for pagagraphs, pages and images <br/> Optional if passing a whole HTML document
         * @param {string} align - Defines the content's alignemnt <br/> Accepts : "left", "right", "center", "justify", "top", "bottom", "middle" <br/> Only works for type: "paragraph" <br/> Optional 
         * @param {number} imageWidth - Specify custom width for images <br/> Only works for type: "image" <br/> Optional
         * @param {number} imageHeight - Specify custom height for images <br/> Only works for type: "image" <br/> Optional
         * @memberof toDoc
         */
        doc.createContent = function(type, content, position, align, imageWidth, imageHeight) {
            var cont = "", 
                contentPosition = 0,
                contentType = "";

            var contentObj = {
                "cContent": cont,
                "cPosition": contentPosition,
                "cType" : contentType,
                "cAlign": "",
                "iAlign": "",
                "iWidth": "",
                "iHeight": "",
            };

            switch (type) {
                
                // Create content with an Image
                case "image":
                    var urlRegex = /^(?:(?:https?|ftp):\/\/)?(?:(?!(?:10|127)(?:\.\d{1,3}){3})(?!(?:169\.254|192\.168)(?:\.\d{1,3}){2})(?!172\.(?:1[6-9]|2\d|3[0-1])(?:\.\d{1,3}){2})(?:[1-9]\d?|1\d\d|2[01]\d|22[0-3])(?:\.(?:1?\d{1,2}|2[0-4]\d|25[0-5])){2}(?:\.(?:[1-9]\d?|1\d\d|2[0-4]\d|25[0-4]))|(?:(?:[a-z\u00a1-\uffff0-9]-*)*[a-z\u00a1-\uffff0-9]+)(?:\.(?:[a-z\u00a1-\uffff0-9]-*)*[a-z\u00a1-\uffff0-9]+)*(?:\.(?:[a-z\u00a1-\uffff]{2,})))(?::\d{2,5})?(?:\/\S*)?$/;
                    // Check for valid URL
                    if (urlRegex.test(content)) {
                        var imgUrl = content;
                        getImage(imgUrl, function(imgString){
                            contentObj.cContent = imgString;
                            contentObj.cType = type;
                            // Check for custom alignment
                            if (align && (align === "left" || align === "top" || align === "middle" || align === "left" || align === "right")) {
                                contentObj.iAlign = align;
                            } else {
                                console.error(align + " is an invalid alignment for image, expected 'left', 'top', 'middle', 'left' or 'right'");
                                return;
                            }

                            // Check for custom width
                            if (imageWidth && (Number.isInteger(imageWidth) && imageWidth > 0)) {
                                contentObj.iWidth = imageWidth;
                            } else {
                                console.error(imageWidth + " is an invalid image width, expected type number");
                                return;
                            }

                            // Check for custom height
                            if (imageHeight && (Number.isInteger(imageHeight) && imageHeight > 0)) {
                                contentObj.iHeight = imageHeight;
                            } else {
                                console.error(imageHeight + " is an invalid image height, expected type number");
                                return;
                            }

                        });
                    } else {
                        console.error(content + " is not an valid value for the parameter 'content', expected image url/path");
                        return;
                    }
                    break;
                
                // Create content with Text or HTML
                case "paragraph":
                    if (typeof(content) == "string" && content.length > 0) {
                        contentObj.cContent = content;
                        // Check if content is Text or HTML markup
                        var htmlRegex = /<\/?[a-z][\s\S]*>/i;
                        if (htmlRegex.test(content)) {
                            contentObj.cType = "html";
                        } else {
                            contentObj.cType = type;
                            // Check for custom alignment
                            if (align && (align === "left" || align === "center" || align === "right" || align === "justify")) {
                                contentObj.cAlign = align;
                            } else {
                                console.error(align + " is an invalid alignment for image, expected 'left', 'center', 'right', or 'justify'");
                                return;
                            }
                        }
                    } else {
                        console.error(content + " is not an valid value for the parameter 'content', expected string");
                        return;
                    }
                    break;
                
                // Create page with Text ot HTML
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
                    console.error(contentType + " is an invalid content type, expected 'image', 'string' or 'html'");
                    return;
            }

            // Check for valid and duplicate content position
            if (Number.isInteger(position) && position > 0) {
                if (oData.aContent.length > 0) {
                    oData.aContent.every(function(i){
                        if (i.hasOwnProperty("cPosition")) {
                            if (i.cPosition == position) {
                                console.error("Content with position: " + i.cPosition + " has already been defined, enter new position for " + content);
                                return false;
                            } else {
                                return true;
                            }
                        }
                    });
                }
                contentObj.cPosition = position; // Check for relocaton
            } else {
                console.error(position + " is an invalid content position, expected type number");
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
                content.sort(function(a,b){
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

                        contentString = contentString + oData.paraStart + i.cContent + oData.paraEnd;
                        if (i.cNextLine) {

                        }

                    } else if (i.cType === "page") {

                        contentString = oData.pageBreak + contentString + i.cContent;

                    } else if (i.cType === "image" || i.cType === "html") {

                        contentString = contentString + oData.break + i.cContent + oData.break;

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
            // set image source
            image.src = url;
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

}) (window);