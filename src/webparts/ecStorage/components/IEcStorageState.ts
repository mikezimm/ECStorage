
import { sp, Views, IViews, ISite } from "@pnp/sp/presets/all";
import { IPickedWebBasic, IPickedList, }  from '@mikezimm/npmfunctions/dist/Lists/IListInterfaces';
import { IUser } from '@mikezimm/npmfunctions/dist/Services/Users/IUserInterfaces';

export interface IECStorageList extends IPickedList {

  Created: string;
  ItemCount: number;
  LastItemUserModifiedDate: string;
  Title: string;
  BaseType: number;
  Id: string;
  DocumentTemplateUrl: string;
  LibraryUrl: string;

}

export interface IItemDetail {
  batch: number; //index of the batch in state.batches
  index: number; //index of item in state.batches[batch].items
  id: number;
  value: number | string; //value to highlight/sort for this detail
  created: any;
  modified: any;
  authorId: number;
  editorId: number;
  authorTitle: string;
  editorTitle: string;
  FileLeafRef: string;
  FileRef: string;
  checkedOutId?: number;
  docIcon?: string;
  uniquePerms?: boolean;
  currentUser: boolean;
  size: number;
  sizeMB: number;
  isFolder?: boolean;
  createYr: number;
  modYr: number;
  bucket: string; // yyyy-mm
  createMs: number;
  modMs: number;

}

export interface IBucketSummary {
  title: string;
  count: number;
  size: number;
  sizeGB: number;
  countP: number;
  sizeP: number;
  users: string[];
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

  uniqueRolls: IItemDetail[];
  typeList: string[];
  types: IFileType[];
  duplicateNames: string[];
  duplicates: IDuplicateFile[];
  folders:  IItemDetail[];

}

export interface IDuplicateFile {
  name: string;
  type: string;
  count: number;
  size: number;
  sizeGB: number;
  items: IItemDetail[];
  locations: string[];
  sizes: number[];
  createdMs: number[];
  modifiedMs: number[];
}

export interface IFileType {
  type: string;
  count: number;
  size: number;
  sizeGB: number;
  items: IItemDetail[];
  sizes: number[];
  createdMs: number[];
  modifiedMs: number[];

}

export interface IUserRanks {
  createSizeRank: number[]; //Array of user index's in the AllUsers array based on this metric.
  createCountRank: number[]; //Array of user index's in the AllUsers array based on this metric.
  modifySizeRank: number[]; //Array of user index's in the AllUsers array based on this metric.
  modifyCountRank: number[]; //Array of user index's in the AllUsers array based on this metric.
}

export interface ITypeRanks {
  sizeRank: number[]; //Array of user index's in the AllUsers array based on this metric.
  countRank: number[]; //Array of user index's in the AllUsers array based on this metric.
}

export interface IDupRanks {
  sizeRank: number[]; //Array of user index's in the AllUsers array based on this metric.
  countRank: number[]; //Array of user index's in the AllUsers array based on this metric.
}

export interface IFolderRanks {
  sizeRank: number[]; //Array of user index's in the AllUsers array based on this metric.
  countRank: number[]; //Array of user index's in the AllUsers array based on this metric.
}

//IBatchData, ILargeFiles, IUserFiles, IOldFiles
export interface IBatchData {
  count: number;
  size: number;
  sizeGB: number;
  large: ILargeFiles;
  oldCreated: IOldFiles;
  oldModified: IOldFiles;
  currentUser: IUserSummary;
  folders:  IItemDetail[];
  folderRanks: IFolderRanks;
  creatorIds: number[];
  editorIds: number[];
  allUsersIds: number[];
  allUsers: IUserSummary[];
  uniqueRolls: IItemDetail[];
  typeList: string[];
  types: IFileType[];
  typeRanks: ITypeRanks;
  duplicateNames: string[];
  duplicates: IDuplicateFile[];
  duplicateRanks: IDupRanks;
  userRanks: IUserRanks;
}

export interface IECStorageBatch {
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

export interface IECStorageFilter {
  startDate: any;
  endDate: any;
  minSize: number;
  maxSize: number;
}



export interface IEcStorageState {

  theSite: ISite;
  pickedWeb : IPickedWebBasic;
  pickedList? : IECStorageList;

  currentUser: IUser;

  parentWeb: string;
  listTitle: string;

  isCurrentWeb: boolean;

  isLoaded: boolean;
  isLoading: boolean;

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

  batches: IECStorageBatch[];
  batchData: IBatchData;

}
