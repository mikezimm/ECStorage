
import { sp, Views, IViews, ISite } from "@pnp/sp/presets/all";
import { IPickedWebBasic, IPickedList, }  from '@mikezimm/npmfunctions/dist/Lists/IListInterfaces';
import { IUser } from '@mikezimm/npmfunctions/dist/Services/Users/IUserInterfaces';

import { IExStorageState, IEXStorageList, IEXStorageBatch, IBatchData, IUserSummary, IItemDetail, IFileType } from '../../IExStorageState';

export interface IExTypesState {

  isLoaded: boolean;
  isLoading: boolean;

  showPane: boolean;
  errorMessage: string;
  hasError: boolean;

  showItems: boolean;
  selectedType: IFileType;

  items: IItemDetail[];

  minYear: number;
  maxYear: number;

  rankSlider: number;
  textSearch: string;

  fetchSlider: number;
  fetchTotal: number;
  fetchCount: number;
  showProgress: boolean;
  fetchPerComp: number;
  fetchLabel: string;

}
