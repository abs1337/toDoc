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
            aHeaderLeft: [],
            aHeaderRight: [],
            aHeaderCenter: [],
            // Footer
            aFooterLeft: [],
            aFooterRight: [],
            aFooterCenter: [],

            // Page / Paragraph
            break: "<br/>",
            aContent: [],
            pageBreak: "<br clear='all' style='page-break-before:always' />",
            paraStart: "<p>",
            paraEnd: "</p>"
        };

        /** 
         * Creates a Header or Footer Section in the document
         * @public
         * @param {string} sectionType - Defines whether the section is a header or footer <br/> Accepts : "header", "footer" <br/> Required
         * @param {string} contentType - Defines whether the section contains an image, text or HTML markup <br/> Accepts : "image", "text", "html" <br/> Required
         * @param {string} content - Defines the sections's content <br/> Accepts : stringified image URLs, stringified text, stringified HTML markup <br/> Required
         * @param {string} alignment - Defines the section content's alignemnt <br/> Accepts : "left", "center", "right" <br/> Required
         * @param {number} position - Specifies the content's position in the section <br/> Accepts : 1++ <br/> Default value :  0 <br/> Optional
         * @param {boolean} nextLine - Specifes whether the section's content should start in a new line <br/> Accepts : true, false <br/> Default value : false <br/> Optional
         * @memberof toDoc
         */
        doc.createSection = function(sectionType, contentType, content, alignment, position, nextLine) {

            var sectionCont = "";
            var sectionContPosition = 0;

            var sectionObj = {
                "sectionContent": sectionCont,
                "sectionPosition": sectionContPosition,
                "sectionNextLine": false,
            };

            switch (contentType) {
                // Create Header Image
                case "image":
                    var urlRegex = /^(?:(?:https?|ftp):\/\/)?(?:(?!(?:10|127)(?:\.\d{1,3}){3})(?!(?:169\.254|192\.168)(?:\.\d{1,3}){2})(?!172\.(?:1[6-9]|2\d|3[0-1])(?:\.\d{1,3}){2})(?:[1-9]\d?|1\d\d|2[01]\d|22[0-3])(?:\.(?:1?\d{1,2}|2[0-4]\d|25[0-5])){2}(?:\.(?:[1-9]\d?|1\d\d|2[0-4]\d|25[0-4]))|(?:(?:[a-z\u00a1-\uffff0-9]-*)*[a-z\u00a1-\uffff0-9]+)(?:\.(?:[a-z\u00a1-\uffff0-9]-*)*[a-z\u00a1-\uffff0-9]+)*(?:\.(?:[a-z\u00a1-\uffff]{2,})))(?::\d{2,5})?(?:\/\S*)?$/;
                    // Check for valid URL
                    if (urlRegex.test(content)) {
                        var imgUrl = content;
                        getImage(imgUrl, function(imgString){
                            sectionObj.sectionContent = imgString
                        });
                    } else {
                        console.error(content + "is not an valid value for the parameter 'content', expected image url/path");
                    }
                    break;
                // Create Header Text
                case "text":
                    if (typeof(content) == "string" && content.length > 0) {
                        sectionObj.sectionContent = content;
                    } else {
                        console.error(content + "is not an valid value for the parameter 'content', expected string");
                    }
                    break;
                // Create Header HTML Markup
                case "html":
                    var htmlRegex = /<\/?[a-z][\s\S]*>/i;
                    // Check for valid HTML
                    if (htmlRegex.test(content)) {
                        sectionObj.sectionContent = content;
                    } else {
                        console.error("Invalid HTML markup");
                    }
                    break;
                // Error
                default:
                    console.error(contentType + " is an invalid content type, expected 'image', 'string' or 'html'");
                    return;
            }

            // Check for position
            if (Number.isInteger(position) && position > 0) {
                sectionObj.sectionPosition = position;
            } else {
                console.error(position + " is an invalid header position, expected number");
            }

            // Check for Next line
            if (typeof(nextLine) == "boolean") {
                sectionObj.sectionNextLine = nextLine;
            } else {
                console.error(nextLine + " is an invalid nextLine flag, expected true or false, defaulting to false");
            }

            // Check Header of Footer and Set Alignment
            if (typeof(sectionType) == "string" && sectionType.length > 0) {
                if (sectionType === "header") {
                    if (typeof(alignment) == "string" && alignment.length > 0) {
                        if (alignment === "left") {
                            oData.aHeaderLeft.push(sectionObj);
                        } else if (alignment === "center") {
                            oData.aHeaderCenter.push(sectionObj);
                        } else if (alignment === "right") {
                            oData.aHeaderRight.push(sectionObj);
                        } else {
                            console.error(alignment + "is an invalid alignment, expected 'left', 'center' or 'right'");
                        }
                    }
                } else if (sectionType === "footer") {
                    if (typeof(alignment) == "string" && alignment.length > 0) {
                        if (alignment === "left") {
                            oData.aFooterLeft.push(sectionObj);
                        } else if (alignment === "center") {
                            oData.aFooterCenter.push(sectionObj);
                        } else if (alignment === "right") {
                            oData.aFooterRight.push(sectionObj);
                        } else {
                            console.error(alignment + "is an invalid alignment, expected 'left', 'center' or 'right'");
                        }
                    }
                } else {
                    console.error(sectionType + " is an invalid section type, expected 'header' or 'footer'");
                }
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

            // Stich header elements together
            if (sections.length > 0) {
                sections.forEach(function(i) {
                    if (i.sectionNextLine) {
                        sectionString = oData.break + i.sectionContent + sectionString;
                    } else {
                        sectionString = i.sectionContent + sectionString;
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
         * @param {boolean} nextLine - Specifes whether the content should start in a new line <br/> Accepts : true, false <br/> Default value : false <br/> Optional
         * @memberof toDoc
         */
        doc.createContent = function(type, content, position, nextLine) {
            var cont = "", 
                contentPosition = 0,
                contentType = "",
                contentNextLine = false;

            var contentObj = {
                "cContent": cont,
                "cPosition": contentPosition,
                "cType" : contentType,
                "cNextLine": contentNextLine
            };

            switch (type) {
                // Create Header Image
                case "image":
                    var urlRegex = /^(?:(?:https?|ftp):\/\/)?(?:(?!(?:10|127)(?:\.\d{1,3}){3})(?!(?:169\.254|192\.168)(?:\.\d{1,3}){2})(?!172\.(?:1[6-9]|2\d|3[0-1])(?:\.\d{1,3}){2})(?:[1-9]\d?|1\d\d|2[01]\d|22[0-3])(?:\.(?:1?\d{1,2}|2[0-4]\d|25[0-5])){2}(?:\.(?:[1-9]\d?|1\d\d|2[0-4]\d|25[0-4]))|(?:(?:[a-z\u00a1-\uffff0-9]-*)*[a-z\u00a1-\uffff0-9]+)(?:\.(?:[a-z\u00a1-\uffff0-9]-*)*[a-z\u00a1-\uffff0-9]+)*(?:\.(?:[a-z\u00a1-\uffff]{2,})))(?::\d{2,5})?(?:\/\S*)?$/;
                    // Check for valid URL
                    if (urlRegex.test(content)) {
                        var imgUrl = content;
                        getImage(imgUrl, function(imgString){
                            contentObj.cContent = imgString;
                            contentObj.cType = type;
                        });
                    } else {
                        console.error(content + "is not an valid value for the parameter 'content', expected image url/path");
                        return;
                    }
                    break;
                // Create Header Text
                case "paragraph":
                    if (typeof(content) == "string" && content.length > 0) {
                        contentObj.cContent = content;

                        // Check for HTML markup
                        var htmlRegex = /<\/?[a-z][\s\S]*>/i;
                        if (htmlRegex.test(content)) {
                            contentObj.cType = "html";
                        } else {
                            contentObj.cType = type;
                        }

                    } else {
                        console.error(content + "is not an valid value for the parameter 'content', expected string");
                        return;
                    }
                    break;
                // Create  Markup
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
                console.error(position + "is an invalid content position, expected type number");
                return;
            }
            
            // Check for Next line
            if (typeof(nextLine) == "boolean") {
                contentObj.cNextLine = nextLine;
            } else {
                console.error(nextLine + " is an invalid nextLine flag, expected true or false, defaulting to false");
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

                    } else if (i.cType === "page") {

                        contentString = oData.pageBreak + contentString + i.cContent;

                    } else if (i.cType === "image") {

                        contentString = contentString + oData.break + i.cContent + oData.break;
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
         * @param {string} url - URL or path of the image.
         * @returns {string} Returns a stringified markup of the image.
         * @memberof toDoc
         */
        var getImage = function(url, callBack) {
            var image = new Image();
            image.crossOrigin = "anonymous"

            // run only after image loads
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

                    // Return data URL in HTML tags
                    if (callBack) {
                        callBack("<img src='" + imgAsDataURL + "' crossOrigin='anonymous'></img>")
                    } else {
                        return "<img src='" + imgAsDataURL + "' crossOrigin='anonymous'></img>";
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
                                                "<tr>"+
                                                    // Header Left Align
                                                    "<td width='30%'>"+
                                                        getSection(oData.aHeaderLeft) +
                                                    "</td>"+
                                                    // Header Center Align
                                                    "<td align='center' width='40%'>"+
                                                        getSection(oData.aHeaderCenter) +
                                                    "</td>"+
                                                    // Header Right Align
                                                    "<td align='right' width='30%'>"+
                                                        getSection(oData.aHeaderRight) +
                                                    "</td>"+
                                                "</tr>"+
                                            "</table>"+
                                        "</p>"+
                                    "</div>"+
                                "</td>"+
                                "<td>"+
                                    //Footer
                                    "<div style='mso-element:footer' id='f1'>"+
                                        "<p class='MsoFooter'>"+
                                            "<table width='100%' border='0' cellspacing='0' cellpadding='0'>"+
                                                "<tr>"+
                                                    // Footer Left Align
                                                    "<td width='30%'>"+
                                                        getSection(oData.aFooterLeft) +
                                                    "</td>"+
                                                    // Footer Center Align
                                                    "<td align='center' width='40%'>"+
                                                        getSection(oData.aFooterCenter) +
                                                    "</td>"+
                                                    // Footer Right Align
                                                    "<td align='right' width='30%'>"+
                                                        getSection(oData.aFooterRight) +
                                                        "Page <span style='mso-field-code: PAGE '></span> of <span style='mso-field-code: NUMPAGES '></span>"+
                                                    "</td>"+
                                                "</tr>"+
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
        };

        return doc;
    }

    // Set Lib globally and save to window
    if (typeof(window.toDoc) === 'undefined') {
        window.toDoc = toDoc();
    }

}) (window);