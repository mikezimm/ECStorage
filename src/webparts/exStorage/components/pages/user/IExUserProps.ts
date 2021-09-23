
import { sp, Views, IViews, ISite } from "@pnp/sp/presets/all";

import { WebPartContext } from '@microsoft/sp-webpart-base';

import { PageContext } from '@microsoft/sp-page-context';

import { escape } from '@microsoft/sp-lodash-subset';

import { IPickedWebBasic, IPickedList, }  from '@mikezimm/npmfunctions/dist/Lists/IListInterfaces';

import { IUser } from '@mikezimm/npmfunctions/dist/Services/Users/IUserInterfaces';

import { IExStorageState, IEXStorageList, IEXStorageBatch, IBatchData, IUserSummary } from '../../IExStorageState';

import { IDataOptions, IUiOptions } from '../../IExStorageProps';

export interface IExUserProps {

      // 0 - Context
      wpContext: WebPartContext;
      pageContext: PageContext;

      WebpartElement: HTMLElement;   //Size courtesy of https://www.netwoven.com/2018/11/13/resizing-of-spfx-react-web-parts-in-different-scenarios/

        // 1 - Analytics options
      // useListAnalytics: boolean;
      // analyticsWeb: string;
      // analyticsList: string;
      tenant: string;
  
      theSite: ISite;
      pickedWeb : IPickedWebBasic;
      pickedList? : IEXStorageList;

      isLoaded: boolean;
  
      currentUser: IUser;
      isCurrentUser: boolean;
  
      userSummary:  IUserSummary;

      batches: IEXStorageBatch[];
      batchData: IBatchData;

      dataOptions: IDataOptions;
      uiOptions: IUiOptions;

      WebpartHeight: number;
      WebpartWidth: number;

}
