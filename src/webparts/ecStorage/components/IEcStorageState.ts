
import { sp, Views, IViews, ISite } from "@pnp/sp/presets/all";
import { IPickedWebBasic, IPickedList, }  from '@mikezimm/npmfunctions/dist/Lists/IListInterfaces';
import { IUser } from '@mikezimm/npmfunctions/dist/Services/Users/IUserInterfaces';

export interface IEcStorageState {

  theSite: ISite;
  pickedWeb : IPickedWebBasic;
  pickedList? : IPickedList;

  currentUser: IUser;
  
  parentWeb: string;
  listTitle: string;

  isCurrentWeb: boolean;

  allLoaded: boolean;

  allowRailsOff: boolean;  //property that determines if the related toggle is visible or not

  showPane: boolean;

  stateError?: any[];

}
