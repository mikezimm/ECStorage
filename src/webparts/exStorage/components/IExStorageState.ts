
import { sp, Views, IViews, ISite } from "@pnp/sp/presets/all";
import { IPickedWebBasic, IPickedList, }  from '@mikezimm/npmfunctions/dist/Lists/IListInterfaces';
import { IUser } from '@mikezimm/npmfunctions/dist/Services/Users/IUserInterfaces';

import { IZLoadAnalytics, IZSentAnalytics, } from '@mikezimm/npmfunctions/dist/Services/Analytics/interfaces';

import { IGridColumns } from './pages/GridCharts/IGridchartsProps';

export interface IEXStorageList extends IPickedList {

  Created: string;
  ItemCount: number;
  LastItemUserModifiedDate: string;
  Title: string;
  BaseType: number;
  Id: string;
  DocumentTemplateUrl: string;
  LibraryUrl: string;
  EntityTypeName: string;
  Hidden: boolean;
  minYear: number;
  maxYear: number;

}

export interface IItemDetail {
  batch: number; //index of the batch in state.batches
  index: number; //index of item in state.batches[batch].items
  id: number;
  value: number | string; //value to highlight/sort for this detail
  Created: any; //This is the actual item Created property.
  created: any;
  Modified: any; //This is the actual item Modified property
  modified: any;
  authorId: number;
  editorId: number;
  authorTitle: string;
  editorTitle: string;
  authorName: string;
  editorName: string;
  parentFolder: string;

  localFolder: string;  //localFolder is the folder Url with the site and library removed... just showing \foldername\subfoldername\

  FileLeafRef: string;
  FileRef: string;
  checkedOutId?: number;
  docIcon?: string;  
  iconName: string;
  iconColor: string;
  iconTitle: string;
  uniquePerms?: boolean;
  
  currentUser: boolean;
  createYr: number;
  modYr: number;
  bucket: string; // yyyy-mm
  createMs: number;
  modMs: number;
  ContentTypeId: string;
  ContentTypeName: string;
  ServerRedirectedEmbedUrl: string; //This property is used to open files correctly... including Word and Excel in the browser

  size: number;
  sizeMB: number;
  sizeLabel: string;

  version: number;
  versionlabel: string;

  isFolder?: boolean;

  MediaServiceAutoTags?: string;
  MediaServiceLocation?: string;
  MediaServiceOCR?: string;
  MediaServiceKeyPoints?: string;
  MediaLengthInSeconds?: string;
  isMedia: boolean;
  whichWasFirst: 'created' | 'modfied';
  whichWasFirstDays: string;

}

export interface IFolderDetail extends IItemDetail {
  directCount: number; //Only next direct children, not their descendants
  directSize: number; //Only next direct children, not their descendants
  directItems: IItemDetail[]; //Only next direct children, not their descendants
  directSizes: number[]; 
  totalCount: number; //Total count including all descendants
  totalSize: number; //Total size including all descendants
  otherItems: IItemDetail[];  //Items in folders below this folder
}


export interface IBucketSummary {
  title: string;
  count: number;
  size: number;
  sizeGB: number;
  sizeLabel: string;
  countP: number;
  sizeP: number;
  ranges: {
    firstCreateMs: any;
    lastCreateMs: any;
    firstModifiedMs: any;
    lastModifiedMs: any;
    createRange: string;
    modifyRange: string;
    firstAllMs: any;
    lastAllMs: any;
    rangeAll: string;
  };
  sizeToCountRatio: number;  //Ratio of sizeP over countP.  Like 75% of all storage is filled by 5% of files ( 75/5 = 15 : 1 )
  userTitles: string[];
  userIds: number[];
}

export interface ILargeFiles {
  GT10G: IItemDetail[];
  GT01G: IItemDetail[];
  GT100M: IItemDetail[];
  GT10M: IItemDetail[];
  summary: IBucketSummary;

}

export interface IOldFiles {
  Age5Yr: IItemDetail[];
  Age4Yr: IItemDetail[];
  Age3Yr: IItemDetail[];
  Age2Yr: IItemDetail[];
  Age1Yr: IItemDetail[];
  summary: IBucketSummary;

}


// export interface IUserFiles {
//   items:  IItemDetail[];
//   large: ILargeFiles;
//   oldCreated: IOldFiles;
//   oldModified: IOldFiles;
//   count: number;
//   size: number;
//   sizeGB: number;
//   summary: IBucketSummary;
// }

export interface IUserSummary {
  userId: number;
  userTitle: string;
  userFirst: any;
  userLast: any;

  folderCreateCount: number;

  createCount: number;
  createSizes: number[];
  createTotalSize: number;
  createTotalSizeLabel: string;
  createTotalSizeGB: number;
  createSizeRank: number;
  createCountRank: number;
  oldCreated: IOldFiles;

  modifyCount: number;
  modifiedSizes: number[];
  modifyTotalSize: number;
  modifyTotalSizeLabel: string;
  modifyTotalSizeGB: number;
  modifySizeRank: number;
  modifyCountRank: number;

  oldModified: IOldFiles;

  summary: IBucketSummary;

  large: ILargeFiles;
  items: IItemDetail[];

  folderInfo: IFolderInfo;

  uniqueInfo: IUniqueInfo;

  typesInfo: ITypeInfo;

  duplicateInfo: IDuplicateInfo;

}

export interface IDuplicateFile {
  name: string;
  type: string;

  iconName: string;
  iconColor: string;
  iconTitle: string;
  
  // These are already in IBucketSummary
  // title: string;
  // count: number;
  // size: number;
  // sizeGB: number;
  // sizeLabel: string;
  // countP: number;
  // sizeP: number;
  // sizeToCountRatio: number;  //Ratio of sizeP over countP.  Like 75% of all storage is filled by 5% of files ( 75/5 = 15 : 1 )
  // userTitles: string[];
  // userIds: number[];
  count: number;
  size: number;
  items: IItemDetail[];
  locations: string[];
  sizes: number[];
  createdMs: number[];
  modifiedMs: number[];
  summary: IBucketSummary;
  isMedia?: boolean;
  FileLeafRef: string;
}

export interface IFileType {
  type: string;
  iconName: string;
  iconColor: string;
  iconTitle: string;
  count: number;
  size: number;
  sizeGB: number;
  sizeP: number;
  countP: number;
  sizeToCountRatio: number;  //Ratio of sizeP over countP.  Like 75% of all storage is filled by 5% of files ( 75/5 = 15 : 1 )
  sizeLabel: string;
  avgSize: number;
  maxSize: number;
  avgSizeLabel: string;
  maxSizeLabel: string;
  items: IItemDetail[];
  sizes: number[];
  createdMs: number[];
  modifiedMs: number[];

}

export interface IUserInfo {
  
  currentUser: IUserSummary;

  count: number;

  creatorIds: number[];
  editorIds: number[];
  allUsersIds: number[];
  allUsers: IUserSummary[];

  createSizeRank: number[]; //Array of user index's in the AllUsers array based on this metric.
  createCountRank: number[]; //Array of user index's in the AllUsers array based on this metric.
  modifySizeRank: number[]; //Array of user index's in the AllUsers array based on this metric.
  modifyCountRank: number[]; //Array of user index's in the AllUsers array based on this metric.
}

export interface ITypeInfo {
  count: number;
  typeList: string[];
  types: IFileType[];
  sizeRank: number[]; //Array of user index's in the AllUsers array based on this metric.
  countRank: number[]; //Array of user index's in the AllUsers array based on this metric.
}

export interface IDuplicateInfo {
  count: number;
  duplicateNames: string[];
  duplicates: IDuplicateFile[];
  sizeRank: number[]; //Array of user index's in the AllUsers array based on this metric.
  countRank: number[]; //Array of user index's in the AllUsers array based on this metric.
  summary: IBucketSummary;
}

export interface IFolderInfo {
  count: number;
  folderRefs: string[];
  folders:  IFolderDetail[];
  sizeRank: number[]; //Array of user index's in the AllUsers array based on this metric.
  countRank: number[]; //Array of user index's in the AllUsers array based on this metric.
}

export interface IUniqueInfo {
  count: number;
  uniqueRolls: IItemDetail[];
}

export type IAllItemTypes = IFolderDetail | IItemDetail;
//IBatchData, ILargeFiles, IUserFiles, IOldFiles

export interface IBatchData {
  totalCount: number;
  count: number;
  size: number;
  sizeGB: number;
  sizeLabel: string;

  large: ILargeFiles;

  oldCreated: IOldFiles;
  oldModified: IOldFiles;

  folderInfo: IFolderInfo;

  userInfo: IUserInfo;

  uniqueInfo: IUniqueInfo;

  typesInfo: ITypeInfo;

  duplicateInfo: IDuplicateInfo;

  items: IAllItemTypes[];

}


export interface IEXStorageBatch {
  index: number;  //Should just be the index of the batch in the batches array
  start: number;
  end: number;
  duration: number;
  msPerItem: number;
  count: number;
  errMessage: string;
  id: string;
  items: any[];
  hasNext: boolean;
  firstCreated: Date;
  lastCreated: Date;
}

export interface IEXStorageFilter {
  startDate: any;
  endDate: any;
  minSize: number;
  maxSize: number;
}

export interface IExStorageState {

  theSite: ISite;
  pickedWeb : IPickedWebBasic;
  pickedList? : IEXStorageList;
  pickLists : IEXStorageList[];

  currentUser: IUser;

  parentWeb: string;
  listTitle: string;

  isCurrentWeb: boolean;

  isLoaded: boolean;
  isLoading: boolean;
  showBegin: boolean;

  allowRailsOff: boolean;  //property that determines if the related toggle is visible or not

  showPane: boolean;
  showUser: number;

  stateError?: any[];
  errorMessage: string;
  hasError: boolean;

  items: any[];

  minYear: number;
  maxYear: number;
  yearSlider: number;

  rankSlider: number;
  userSearch: string;

  fetchSlider: number;
  fetchTotal: number;
  fetchCount: number;
  showProgress: boolean;
  fetchPerComp: number;
  fetchLabel: string;

  batches: IEXStorageBatch[];
  batchData: IBatchData;
  mainGridColumns: IGridColumns;
  
  dropDownLabels: any[];
  dropDownIndex: number;
  dropDownText: string;

  loadProperties: IZLoadAnalytics;

  refreshId: string; //used to trigger redraw of grid

}
