
import { sp, Views, IViews, ISite } from "@pnp/sp/presets/all";
import { IPickedWebBasic, IPickedList, }  from '@mikezimm/npmfunctions/dist/Lists/IListInterfaces';
import { IUser } from '@mikezimm/npmfunctions/dist/Services/Users/IUserInterfaces';

import { IZLoadAnalytics, IZSentAnalytics, } from '@mikezimm/npmfunctions/dist/Services/Analytics/interfaces';

import { IGridColumns } from './pages/GridCharts/IGridchartsProps';

import { IItemSharingInfo, ISharingEvent, ISharedWithUser } from './Sharing/ISharingInterface';

export type IItemType = 'Items' | 'Duplicates' | 'Shared' | 'CheckedOut' ;

export interface IIconArray {
  iconTitle: string;
  iconName: string;
  iconColor: string;
  iconSearch: string;
  other?: any; //Can be used as anything as needed such as a way to sort - ie count of items with this icon
  sort1?: any; //Can be used as anything as needed such as a way to sort - ie count of items with this icon
  sort2?: any; //Can be used as anything as needed such as a way to sort - ie count of items with this icon
}

export interface IEXStorageList extends IPickedList {

  Created: string;
  ItemCount: number;
  LastItemUserModifiedDate: string;
  Title: string;
  DropDownLabel: string;
  BaseType: number;
  Id: string;
  DocumentTemplateUrl: string;
  LibraryUrl: string;
  EntityTypeName: string;
  Hidden: boolean;
  minYear: number;
  maxYear: number;

}

/**
 * 0 are draft, 1 are one version, 100 are 1 to 100, 1000 are 101 to 1000
 */
export type IVersionBucket = 0 | 1 | 1.1 | 100 | 500;

export type IVersionBucketLabel = 'IsDraft' | '1.0' | '>1.0' | '>=100' | '>=500' | 'IsMinor' | 'CheckedOut' | 'CheckedOutToYou';

export type IKnownFileTypes = 'Type:Excel' | 'Type:Word' | 'Type:PowerPoint' | 'Type:Text' | 'Type:pdf' | 'Type:OneNote' | 'Type:Outlook' | 'Type:Zipped' | 'Type:Movie' | 'Type:Image' | 'Type:Dwg' | 'Type:File' ;

//'MediaServiceAutoTags','MediaServiceLocation','MediaServiceOCR','MediaServiceKeyPoints','MediaLengthInSeconds'
export type IKnownMeta = 'Type:Excel' | 'Type:Word' | 'Type:PowerPoint' | 'Type:Text' | 'Type:pdf' | 'Type:OneNote' | 'Type:Outlook' | 'Type:Zipped' | 'Type:Movie' | 'Type:Image' | 'Type:Dwg' | 'Type:File' | 'WasShared' |  'UniquePermissions' | 'Type:Folder' | 'IsDraft' | 'IsMinor' | 'IsMajor' | 'SingleVerion' | 'MediaServiceAutoTags' | 'MediaServiceLocation' | 'MediaServiceOCR' | 'MediaServiceKeyPoints' | 'MediaLengthInSeconds' | '' |
'1.0' | '>1.0' | '>=100' | '>=500' | 'CheckedOut' | 'CheckedOutToYou' ;

export interface IFileVersionInfo {
    number: number;
    string: string;
    bucket: IVersionBucket;
    bucketLabel: IVersionBucketLabel;
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
  authorShared: string;
  editorName: string;
  editorShared: string;
  parentFolder: string;

  localFolder: string;  //localFolder is the folder Url with the site and library removed... just showing \foldername\subfoldername\

  FileLeafRef: string;
  FileRef: string;
  checkedOutId: number;
  checkedOutCurrentUser: boolean;
  docIcon?: IKnownMeta;  
  iconName: string;
  iconColor: string;
  iconTitle: string;

  iconSearch: IKnownMeta; //Tried removing this but it caused issues with the auto-create title icons in Items.tsx so I'm adding it back.

  meta: IKnownMeta[];

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

  version: IFileVersionInfo;

  isFolder?: boolean;

  MediaServiceAutoTags?: string;
  MediaServiceLocation?: string;
  MediaServiceOCR?: string;
  MediaServiceKeyPoints?: string;
  MediaLengthInSeconds?: string;
  isMedia: boolean;
  whichWasFirst: 'created' | 'modfied';
  whichWasFirstDays: string;

  itemSharingInfo?: IItemSharingInfo;

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

export type IBucketType = 'Batch' | 'User' | 'Old Files' | 'Large Files' | 'Duplicate Files' | 'Files with Unique Permissions' | 'Folders' | 'File Type' | 'Shared Files' ;

export interface IBucketSummary {
  title: string;
  count: number;
  size: number;
  sizeGB: number;
  sizeLabel: string;
  countP: number;
  sizeP: number;
  bucket: IBucketType;
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
  itemIds: number[];

}

export interface ILargeFiles {
  GT10G: IItemDetail[];
  GT01G: IItemDetail[];
  GT100M: IItemDetail[];
  GT10M: IItemDetail[];
  summary: IBucketSummary;

}

export interface IVersionInfo {
  draft: IItemDetail[];
  one: IItemDetail[];
  GT1: IItemDetail[];
  GT100: IItemDetail[];
  GT500: IItemDetail[];
  checkedOut: IItemDetail[];
  minor: IItemDetail[];

  // summary: IBucketSummary;

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
  sharedName: string;

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
  
  sharingInfo: ISharingInfo;
  
  versionInfo: IVersionInfo;

}

export interface IDuplicateFile {
  name: string;
  type: string;

  iconName: string;
  iconColor: string;
  iconTitle: string;

  iconSearch: IKnownMeta; //Tried removing this but it caused issues with the auto-create title icons in Items.tsx so I'm adding it back.
  meta: IKnownMeta[];
  
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

  avgSize: number;
  maxSize: number;
  avgSizeLabel: string;
  maxSizeLabel: string;

  items: IItemDetail[];
  sizes: number[];
  createdMs: number[];
  modifiedMs: number[];
    
  versionInfo: IVersionInfo;

  summary: IBucketSummary;

}

export interface ISharingInfo {

  sharedItems: IItemDetail[];
  summary: IBucketSummary;

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
  duplicateNames: string[];
  allNames: string[];
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
  uniqueRolls: IItemDetail[];
  summary: IBucketSummary;
}

export type IAllItemTypes = IFolderDetail | IItemDetail;
//IBatchData, ILargeFiles, IUserFiles, IOldFiles

export interface IBatchData {
  totalCount: number;

  large: ILargeFiles;

  oldCreated: IOldFiles;
  oldModified: IOldFiles;

  folderInfo: IFolderInfo;

  userInfo: IUserInfo;

  uniqueInfo: IUniqueInfo;

  typesInfo: ITypeInfo;

  duplicateInfo: IDuplicateInfo;

  items: IAllItemTypes[];

  significance: number; // % of all items returned
  isSignificant: boolean;

  summary: IBucketSummary;

  sharingInfo: ISharingInfo;

  versionInfo: IVersionInfo;

  analytics: {
    fetchMs: number,
    analyzeMs: number,
    fetchTime: any,
    fetchDuration: string,
    analyzeDuration: string,
    count: number,
    msPerFetch: number,
    msPerAnalyze: number,

  };

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
