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
            orientation: "portrait",
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
         * Store page, header, footer data
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
            // Page
            break: "<br/>",
            aPages: [],
            pagePre: "<br/>",
            pagePost: "<br clear='all' style='page-break-before:always' />",
        };

        /** 
         * Creates a Header or a Footer object for the document
         * @public
         * @param {string} headerFooter - Defines whether the object is header or footer <br/> Accepts : "header" | "footer" <br/> Required
         * @param {string} type - Defines whether the object contains an image, text or HTML markup <br/> Accepts : "image" | "text" | "html" <br/> Required
         * @param {string} content - Defines the object's content <br/> Accepts : stringified image URLs | stringified text | stringified HTML markup <br/> Required
         * @param {string} alignment - Defines the object's alignemnt <br/> Accepts : "left" | "center"| "right" <br/> Required
         * @param {number} position - Specifies the position of the object <br/> Accepts : 1++ <br/> Default value :  0 <br/> Optional
         * @param {boolean} nextLine - Specifes whether the object should start in a new line <br/> Accepts : true | false <br/> Default value : false <br/> Optional
         * @memberof toDoc
         */
        doc.createHeaderFooter = function(headerFooter, type, content, alignment, position, nextLine) {

            var headerCont = "";
            var headerContPosition = 0;

            var headerFooterObj = {
                "HFContent": headerCont,
                "HFPosition": headerContPosition,
                "HFNextLine": false,
            };

            switch (type) {
                // Create Header Image
                case "image":
                    var urlRegex = /^(?:(?:https?|ftp):\/\/)?(?:(?!(?:10|127)(?:\.\d{1,3}){3})(?!(?:169\.254|192\.168)(?:\.\d{1,3}){2})(?!172\.(?:1[6-9]|2\d|3[0-1])(?:\.\d{1,3}){2})(?:[1-9]\d?|1\d\d|2[01]\d|22[0-3])(?:\.(?:1?\d{1,2}|2[0-4]\d|25[0-5])){2}(?:\.(?:[1-9]\d?|1\d\d|2[0-4]\d|25[0-4]))|(?:(?:[a-z\u00a1-\uffff0-9]-*)*[a-z\u00a1-\uffff0-9]+)(?:\.(?:[a-z\u00a1-\uffff0-9]-*)*[a-z\u00a1-\uffff0-9]+)*(?:\.(?:[a-z\u00a1-\uffff]{2,})))(?::\d{2,5})?(?:\/\S*)?$/;
                    // Check for valid URL
                    if (urlRegex.test(content)) {
                        var imgUrl = content;
                        headerFooterObj.HFContent = getImage(imgUrl);
                    } else {
                        console.log("invalid image url or path");
                    }
                    break;

                    // Create Header Text
                case "text":
                    if (typeof(content) == "string" && content.length > 0) {
                        headerFooterObj.HFContent = content;
                    } else {
                        console.log("invalid text string");
                    }
                    break;

                    // Create Header HTML Markup
                case "html":
                    var htmlRegex = /<\/?[a-z][\s\S]*>/i;
                    // Check for valid HTML
                    if (htmlRegex.test(content)) {
                        headerFooterObj.HFContent = content;
                    } else {
                        console.log("invalid html markup");
                    }
                    break;

                    // Error
                default:
                    console.log("invalid header type");
                    return;
            }

            // Check for position
            if (Number.isInteger(position) && position > 0) {
                headerFooterObj.HFPosition = position;
            } else {
                console.log("invalid header position");
            }

            // Check for Next line
            if (typeof(nextLine) == "boolean") {
                headerFooterObj.HFNextLine = nextLine;
            } else {
                console.log("invalid nextLine parameter");
            }

            // Check Header of Footer and Set Alignment
            if (typeof(headerFooter) == "string" && headerFooter.length > 0) {
                if (headerFooter == "header") {
                    if (typeof(nextLine) == "string" && alignment.length > 0) {
                        if (alignment == "left") {
                            oData.aHeaderLeft.push(headerFooterObj);
                        } else if (alignment == "center") {
                            oData.aHeaderCenter.push(headerFooterObj);
                        } else if (alignment == "right") {
                            oData.aHeaderRight.push(headerFooterObj);
                        } else {
                            console.log("invalid alignment parameter");
                        }
                    }
                } else if (headerFooter == "footer") {
                    if (typeof(nextLine) == "string" && alignment.length > 0) {
                        if (alignment == "left") {
                            oData.aFooterLeft.push(headerFooterObj);
                        } else if (alignment == "center") {
                            oData.aFooterCenter.push(headerFooterObj);
                        } else if (alignment == "right") {
                            oData.aFooterRight.push(headerFooterObj);
                        } else {
                            console.log("invalid alignment parameter");
                        }
                    }
                } else {
                    console.log("invalid header/footer parameter");
                }
            }
        };

        /** 
         * Stitch all Header and Footer objects during document generation
         * @private
         * @returns {string} Returns stringified markup of header and footer content
         * @memberof toDoc
         */
        var geHeaderFooter = function(hfArray) {
            var leftHeaders = hfArray;
            var numOfHeaders = leftHeaders.length;
            var hfString = "";

            // Stich header elements together
            if (numOfHeaders > 0) {
                leftHeaders.forEach(function(i) {
                    if (i.HFNextLine) {
                        hfString = oData.break+i.HFContent + hfString;
                    } else {
                        hfString = i.HFContent + hfString;
                    }
                });
            }
            return hfString;
        };


        /** 
         * Creates a Page object for the document
         * @public
         * @param {string} content - Defines the page's content <br/> Accepts : stringified text | stringified HTML markup <br/> Required
         * @param {number} number - Defines the Pages's number <br/> Accepts : 1+= <br/> Default value : 0 <br/> Required for individual pages <br/> Required if creating individual pages <br/> Optional if passing a whole HTML document
         * @memberof toDoc
         */
        doc.createPage = function(content, number) {
            var pageNum = 0;
            var pageCont = "";
            // Check for valid page number
            if (Number.isInteger(number) && number > 0) {
                pageNum = number;
            } else {
                console.log("invalid page number");
                return;
            }
            // Check for valid page content
            if (typeof(content) == "string" && content.length > 0) {
                pageCont = content;
            } else {
                console.log("invalid page content");
                return;
            }

            // Check for existing page numbers
            if (oData.aPages.length > 0) {
                oData.aPages.forEach(function(i){
                    if (i.hasOwnProperty("pNumber")) {
                        console.log("pages with page number exist, following pages must also have page numbers")
                        return;
                    }
                });
            }

            // Create page object and store in pages array
            if (pageNum != 0 && pageCont != "") {
                var pageObj = {
                    "pNumber": pageNum,
                    "pContent": pageCont
                };
                oData.aPages.push(pageObj);
            } else if (!pageNum && pageCont != "") {
                var pageObj = {
                    "pNumber": pageNum,
                };
                oData.aPages.push(pageObj);
            }
        };

        /** 
         * Stitch all Page objects during document generation
         * @private
         * @returns {string} Returns stringified markup of all pages
         * @memberof toDoc
         */
        var getPages = function() {
            var pages = oData.aPages;

            if (pages.length > 0) {
                var pageString = "";
                pages.forEach(function(i, index) {
                    // Check for last page
                    if (i.pNumber == (pages.length-1)) {
                        pageString = pageString + oData.pagePre + i.pContent;
                    } else {
                        pageString = pageString + oData.pagePre + i.pContent + oData.pagePost;
                    }
                });
                return pageString;
            } else {
                return "";
            }
        };

        /** 
         * Creates a Page object for the document
         * @public
         * @param {string} content - Defines the page's content <br/> Accepts : stringified text | stringified HTML markup <br/> Required
         * @param {number} number - Defines the Pages's number <br/> Accepts : 1+= <br/> Default value : 0 <br/> Required for individual pages <br/> Required if creating individual pages <br/> Optional if passing a whole HTML document
         * @memberof toDoc
         */
        doc.createParagraph = function() {

        };


        /** 
         * Creates and stores an image as a local dataURL from image's URL
         * @private
         * @param {string} url - URL or path of the image.
         * @returns {string} Returns a stringified markup of the image.
         * @memberof toDoc
         */
        var getImage = function(url) {
            var image = new Image();
            image.src = url;
            if (image != null && (image.width && image.height)) {

                var imgCanvas = document.createElement("canvas"),
                    imgContext = imgCanvas.getContext("2d");

                // Make sure canvas is as big as the picture
                imgCanvas.width = image.width;
                imgCanvas.height = image.height;

                // Draw image into canvas element
                imgContext.drawImage(image, 0, 0, image.width, image.height);

                // Get canvas contents as a data URL
                var imgAsDataURL = imgCanvas.toDataURL("image/png");

                return "<img src='" + imgAsDataURL + "'></img>";
            } else {
                return "";
            }
        };


        /** 
         * Generates and saves a Word document.
         * @public
         * @param {string} filename - Specifies the name of the document <br/> Required
         * @param {object} params - Define custom parameters for the document's layout <br/> Optional
         * @memberof toDoc
         */
        doc.createDocument = function(filename, params) {

            var docParams = {};

            // Validate parameters object
            if (params && !params instanceof Array) {

                Object.keys(params).forEach(function(i) {
                    if (i == "orientation") {
                        if (!params[i] == "portrait" || !params[i] == "landscape") {
                            console.log("invalid page orientation");
                            return;
                        }
                    } else {
                        // Check for number followed by 'cm' or 'in'
                        var unitRegex = /(?:[^\d]\.| |^)((?:\d+\.)?\d+) *cm$|in$/;
                        if (!unitRegex.test(params[i])) {
                            console.log("invalid " + i);
                            return;
                        }
                    }
                });
                docParams = params;
            } else {
                docParams = oSettings;
            }

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
                            getPages() +
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
                                                            geHeaderFooter(oData.aHeaderLeft) +
                                                        "</td>"+
                                                        // Header Center Align
                                                        "<td align='center' width='40%'>"+
                                                            geHeaderFooter(oData.aHeaderCenter) +
                                                        "</td>"+
                                                        // Header Right Align
                                                        "<td align='right' width='30%'>"+
                                                            geHeaderFooter(oData.aHeaderRight) +
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
                                                            geHeaderFooter(oData.aFooterLeft) +
                                                        "</td>"+
                                                        // Footer Center Align
                                                        "<td align='center' width='40%'>"+
                                                            geHeaderFooter(oData.aFooterCenter) +
                                                        "</td>"+
                                                        // Footer Right Align
                                                        "<td align='right' width='30%'>"+
                                                            geHeaderFooter(oData.aFooterRight) +
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

            // Stich all HTMl Markups together and save as a blob
            var html = documentMarkup;

            var blob = new Blob(['\ufeff', html], {
                type: 'application/msword'
            });
            //saveAs(blob, "filename.doc");

            // Get link url
            var url = URL.createObjectURL(blob);

            // Set file name
            if (!filename) {
                filename = filename ? filename + '.doc' : 'document.doc';
            }

            // Create download link element
            var downloadLink = document.createElement("a");
            document.body.appendChild(downloadLink);
            // Check for IE 10/11
            if (navigator.msSaveOrOpenBlob) {
                navigator.msSaveOrOpenBlob(blob, filename);
            } else {
                // Create a link to the file
                downloadLink.href = url;
                // Setting the file name
                downloadLink.download = filename;
                //triggering the function
                downloadLink.click();
            }
            document.body.removeChild(downloadLink);
        };

        return doc;
    }

    // Set lib globally and save to window
    if (typeof(window.toDoc) === 'undefined') {
        window.toDoc = toDoc();
    }
})(window);