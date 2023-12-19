import {
  WebPartContext
} from '@microsoft/sp-webpart-base';
export interface IFileUploadProps {
  digest:string;
  context:WebPartContext;
  listName:string;
  fileTypes:string;
  queryString:string;
  uploadFilesTo:string;
}
