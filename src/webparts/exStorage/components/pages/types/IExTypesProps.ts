
import { sp, Views, IViews, ISite } from "@pnp/sp/presets/all";

import { WebPartContext } from '@microsoft/sp-webpart-base';

import { PageContext } from '@microsoft/sp-page-context';

import { escape } from '@microsoft/sp-lodash-subset';

import { IPickedWebBasic, IPickedList, }  from '@mikezimm/npmfunctions/dist/Lists/IListInterfaces';

import { IUser } from '@mikezimm/npmfunctions/dist/Services/Users/IUserInterfaces';

import { IExStorageState, IEXStorageList, IEXStorageBatch, IBatchData, IUserSummary, ITypeInfo } from '../../IExStorageState';

import { IDataOptions, IUiOptions } from '../../IExStorageProps';

export interface IExTypesProps {

      // 0 - Context
      // wpContext: WebPartContext;
      // pageContext: PageContext;

      // WebpartElement: HTMLElement;   //Size courtesy of https://www.netwoven.com/2018/11/13/resizing-of-spfx-react-web-parts-in-different-scenarios/

      // tenant: string;
  
      theSite: ISite;
      pickedWeb : IPickedWebBasic;
      pickedList? : IEXStorageList;
  
      // currentUser: IUser;
  
      typesInfo:  ITypeInfo;

      batches: IEXStorageBatch[];
      batchData: IBatchData;

      heading: string;

      WebpartHeight: number;
      WebpartWidth: number;

      dataOptions: IDataOptions;
      uiOptions: IUiOptions;

}
