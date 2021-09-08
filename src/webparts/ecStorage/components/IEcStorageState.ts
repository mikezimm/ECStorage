
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

export interface IECStorageBatch {
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

  fetchTotal: number;
  fetchCount: number;
  showProgress: boolean;
  fetchPerComp: number;
  fetchLabel: string;

  batches: IECStorageBatch[];

}
