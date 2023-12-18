"use strict";
var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
var React = require("react");
var sp_core_library_1 = require("@microsoft/sp-core-library");
var react_dropzone_component_1 = require("react-dropzone-component");
var sp_pnp_js_1 = require("sp-pnp-js");
var FileUpload = /** @class */ (function (_super) {
    __extends(FileUpload, _super);
    function FileUpload(props) {
        return _super.call(this, props) || this;
    }
    FileUpload.prototype.render = function () {
        var _context = this.props.context;
        var _listName = this.props.listName;
        var _fileUploadTo = this.props.uploadFilesTo;
        var _queryStringParam = this.props.queryString;
        var queryParameters = new sp_core_library_1.UrlQueryParameterCollection(window.location.href);
        var _itemId = queryParameters.getValue(_queryStringParam);
        var _parent = this;
        var componentConfig = {
            iconFiletypes: this.props.fileTypes.split(','),
            showFiletypeIcon: true,
            postUrl: _context.pageContext.web.absoluteUrl
        };
        var myDropzone;
        var eventHandlers = {
            // This one receives the dropzone object as the first parameter
            // and can be used to additional work with the dropzone.js
            // object
            init: function (dz) {
                myDropzone = dz;
            },
            removedfile: function (file) {
                var web = new sp_pnp_js_1.Web(_context.pageContext.web.absoluteUrl);
                if (_fileUploadTo == "DocumentLibrary") {
                    web.lists.getById(_listName).rootFolder.files.getByName(file.name).delete().then(function (t) {
                        //add your code here if you want to do more after deleting the file
                    });
                }
                else {
                    web.lists.getById(_listName).items.getById(Number(_itemId)).attachmentFiles.deleteMultiple(file.name).then(function (t) {
                        //add your code here if you want to do more after deleting the file
                    });
                }
            },
            processing: function (file) {
                if (_fileUploadTo == "DocumentLibrary")
                    myDropzone.options.url = "".concat(_context.pageContext.web.absoluteUrl, "/_api/web/Lists/getById('").concat(_parent.props.listName, "')/rootfolder/files/add(overwrite=true,url='").concat(file.name, "')");
                else {
                    if (_itemId)
                        myDropzone.options.url = "".concat(_context.pageContext.web.absoluteUrl, "/_api/web/lists/getById('").concat(_parent.props.listName, "')/items(").concat(_itemId, ")/AttachmentFiles/add(FileName='").concat(file.name, "')");
                    else
                        alert('Item not found or query string value is null!');
                }
            },
            sending: function (file, xhr) {
                var _send = xhr.send;
                xhr.send = function () {
                    _send.call(xhr, file);
                };
            },
            error: function (file, error) {
                if (_fileUploadTo != "DocumentLibrary")
                    alert("File '".concat(file.name, "' is already exists, please rename your file or select another file."));
                //if(myDropzone)
                //  myDropzone.removeFile(file);
            }
        };
        var djsConfig = {
            headers: {
                "X-RequestDigest": this.props.digest
            },
            addRemoveLinks: true
        };
        return (React.createElement(react_dropzone_component_1.DropzoneComponent, { eventHandlers: eventHandlers, djsConfig: djsConfig, config: componentConfig },
            React.createElement("div", { className: "dz-message icon ion-upload" }, "Drop files here or click to upload.")));
    };
    return FileUpload;
}(React.Component));
exports.default = FileUpload;
//# sourceMappingURL=FileUpload.js.map