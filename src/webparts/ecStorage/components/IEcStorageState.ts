
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

}

export interface IItemDetail {
  batch: number; //index of the batch in state.batches
  index: number; //index of item in state.batches[batch].items
  id: number;
  value: number | string; //value to highlight/sort for this detail
  created: any;
  modified: any;
  createdId: number;
  modifiedId: number;
  createdTitle: string;
  modifiedTitle: string;
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

export interface ILargeFiles {
  GT10G: IItemDetail[];
  GT01G: IItemDetail[];
  GT100M: IItemDetail[];
  GT10M: IItemDetail[];
}

export interface IOldFiles {
  Age5Yr: IItemDetail[];
  Age4Yr: IItemDetail[];
  Age3Yr: IItemDetail[];
  Age2Yr: IItemDetail[];
  Age1Yr: IItemDetail[];
}

export interface IUserFiles {
  items:  IItemDetail[];
  large: ILargeFiles;
  oldCreate: IOldFiles;
  oldModified: IOldFiles;
}

export interface IUserSummary {
  userId: number;
  userTitle: string;
  userFirst: any;
  userLast: any;
  createCount: number;
  modifyCount: number;
  folderCreateCount: number;
  createTotalSize: number;
  modifyTotalSize: number;
  createSizes: number[];
  modifiedSizes: number[];
}

//IBatchData, ILargeFiles, IUserFiles, IOldFiles
export interface IBatchData {
  large: ILargeFiles;
  oldCreate: IOldFiles;
  oldModified: IOldFiles;
  currentUser: IUserFiles;
  folders:  IItemDetail[];
  creatorIds: number[];
  editorIds: number[];
  allUsersIds: number[];
  allUsers: IUserSummary[];

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

  stateError?: any[];
  errorMessage: string;
  hasError: boolean;

  items: any[];

  minYear: number;
  maxYear: number;
  yearSlider: number;

  fetchSlider: number;
  fetchTotal: number;
  fetchCount: number;
  showProgress: boolean;
  fetchPerComp: number;
  fetchLabel: string;

  batches: IECStorageBatch[];
  batchData: IBatchData;

}
