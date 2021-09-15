
import { sp, Views, IViews, ISite } from "@pnp/sp/presets/all";

import { WebPartContext } from '@microsoft/sp-webpart-base';

import { PageContext } from '@microsoft/sp-page-context';

import { escape } from '@microsoft/sp-lodash-subset';

import { IPickedWebBasic, IPickedList, }  from '@mikezimm/npmfunctions/dist/Lists/IListInterfaces';

import { IUser } from '@mikezimm/npmfunctions/dist/Services/Users/IUserInterfaces';

import { IEcStorageState, IECStorageList, IECStorageBatch, IBatchData, IUserSummary, ITypeInfo } from '../../IEcStorageState';

export interface IEsTypesProps {

      // 0 - Context
      // wpContext: WebPartContext;
      // pageContext: PageContext;

      // WebpartElement: HTMLElement;   //Size courtesy of https://www.netwoven.com/2018/11/13/resizing-of-spfx-react-web-parts-in-different-scenarios/

      // tenant: string;
  
      theSite: ISite;
      pickedWeb : IPickedWebBasic;
      pickedList? : IECStorageList;
  
      // currentUser: IUser;
  
      typesInfo:  ITypeInfo;

      batches: IECStorageBatch[];
      batchData: IBatchData;

      WebpartHeight: number;
      WebpartWidth: number;

}
