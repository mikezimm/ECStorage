
import { sp, Views, IViews, ISite } from "@pnp/sp/presets/all";

import { WebPartContext } from '@microsoft/sp-webpart-base';

import { PageContext } from '@microsoft/sp-page-context';

import { escape } from '@microsoft/sp-lodash-subset';

import { IPickedWebBasic, IPickedList, }  from '@mikezimm/npmfunctions/dist/Lists/IListInterfaces';

import { IUser } from '@mikezimm/npmfunctions/dist/Services/Users/IUserInterfaces';

import { IExStorageState, IEXStorageList, IEXStorageBatch, IBatchData, IUserSummary, ITypeInfo, ILargeFiles, IVersionInfo } from '../../IExStorageState';

import { IDataOptions, IUiOptions } from '../../IExStorageProps';

import {  IGridStyles , IGridColumns} from '../GridCharts/IGridchartsProps';

export interface IExVersionsProps {

      // 0 - Context
      // wpContext: WebPartContext;
      // pageContext: PageContext;

      // WebpartElement: HTMLElement;   //Size courtesy of https://www.netwoven.com/2018/11/13/resizing-of-spfx-react-web-parts-in-different-scenarios/

      // tenant: string;
  
      theSite: ISite;
      pickedWeb : IPickedWebBasic;
      pickedList? : IEXStorageList;
  
      versionInfo: IVersionInfo;
      batchData: IBatchData;

      wpContext: WebPartContext;
      pageContext: PageContext;
      WebpartElement: HTMLElement;   //Size courtesy of https://www.netwoven.com/2018/11/13/resizing-of-spfx-react-web-parts-in-different-scenarios/
      tenant: string;

      WebpartHeight: number;
      WebpartWidth: number;

      heading: string;
      
      dataOptions: IDataOptions;
      uiOptions: IUiOptions;

      columns: IGridColumns;
      gridStyles: IGridStyles;
      refreshId?: string; //used to trigger redraw of grid

}
