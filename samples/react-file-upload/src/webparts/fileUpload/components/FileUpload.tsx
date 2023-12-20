import * as React from 'react';
import * as styles from './FileUpload.module.scss';
import { IFileUploadProps } from './IFileUploadProps';
//import { escape } from '@microsoft/sp-lodash-subset';
import { UrlQueryParameterCollection } from '@microsoft/sp-core-library';
import { DropzoneComponent } from 'react-dropzone-component';
import { Web } from 'sp-pnp-js';
export default class FileUpload extends React.Component<IFileUploadProps, {}> {
  constructor(props: IFileUploadProps) {
    super(props);
  }
  public render(): React.ReactElement<IFileUploadProps> {
    const _context = this.props.context;
    const _listName = this.props.listName;
    const _fileUploadTo = this.props.uploadFilesTo;
    const _queryStringParam = this.props.queryString;
    const queryParameters = new UrlQueryParameterCollection(window.location.href);
    const _itemId = queryParameters.getValue(_queryStringParam);
    // eslint-disable-next-line @typescript-eslint/no-this-alias
    const _parent = this;
    const componentConfig = {
      iconFiletypes: this.props.fileTypes.split(','),
      showFiletypeIcon: true,
      postUrl: _context.pageContext.web.absoluteUrl
    };
    let myDropzone: { options: { url: string; }; };
    const eventHandlers = {
      // This one receives the dropzone object as the first parameter
      // and can be used to additional work with the dropzone.js
      // object
      init: function (dz: { options: { url: string; }; }) {
        myDropzone = dz;
      },
      removedfile: function (file: { name: string; }) {
        const web: Web = new Web(_context.pageContext.web.absoluteUrl);
        if (_fileUploadTo === "DocumentLibrary") {
          // tslint:disable-next-line:no-unsafe-any
          web.lists.getById(_listName).rootFolder.files.getByName(file.name).delete().then((t): void => {
            //add your code here if you want to do more after deleting the file
          }).catch(error => {
            //handle your error here
          });
        }
        else {
          web.lists.getById(_listName).items.getById(Number(_itemId)).attachmentFiles.deleteMultiple(file.name).then((t): void => {
            //add your code here if you want to do more after deleting the file
          }).catch(error => {
            //handle your error here
          });
        }
      },
      processing: function (file: { name: any; }) {
        if (_fileUploadTo === "DocumentLibrary")
          myDropzone.options.url = `${_context.pageContext.web.absoluteUrl}/_api/web/Lists/getById('${_parent.props.listName}')/rootfolder/files/add(overwrite=true,url='${file.name}')`;
        else {
          if (_itemId)
            myDropzone.options.url = `${_context.pageContext.web.absoluteUrl}/_api/web/lists/getById('${_parent.props.listName}')/items(${_itemId})/AttachmentFiles/add(FileName='${file.name}')`;
          else
            alert('Item not found or query string value is null!')
        }
      },
      sending: function (file: any, xhr: { send: () => void; }) {
        const _send = xhr.send;
        xhr.send = function () {
          _send.call(xhr, file);
        };
      },
      error: function (file: { name: any; }, error: any) {
        if (_fileUploadTo !== "DocumentLibrary")
          alert(`File '${file.name}' is already exists, please rename your file or select another file.`);
        //if(myDropzone)
        //  myDropzone.removeFile(file);
      }
    };
    const djsConfig = {
      headers: {
        "X-RequestDigest": this.props.digest
      },
      addRemoveLinks: true,
      acceptedFiles: this.acceptedFilesTypes(this.props.fileTypes),
    };
    return (
      <DropzoneComponent eventHandlers={eventHandlers} djsConfig={djsConfig} config={componentConfig}>
        <div className="dz-message icon ion-upload">Drop files here or click to upload.</div>
      </DropzoneComponent>
    );
  }

  private acceptedFilesTypes(fileTypes: string): string {
    const acceptedFiles = fileTypes.split(',');
    for (let i = 0; i < acceptedFiles.length; i++) {
      if (acceptedFiles[i].lastIndexOf('.', 0) === -1) {
        acceptedFiles[i] = `.${acceptedFiles[i]}`;
      }
    }
    return acceptedFiles.join(',');
  }
}
