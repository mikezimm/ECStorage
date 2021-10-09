
import { sp, Views, IViews, ISite } from "@pnp/sp/presets/all";

import { WebPartContext } from '@microsoft/sp-webpart-base';

import { PageContext } from '@microsoft/sp-page-context';

import { escape } from '@microsoft/sp-lodash-subset';

import { IPickedWebBasic, IPickedList, }  from '@mikezimm/npmfunctions/dist/Lists/IListInterfaces';

import { IUser } from '@mikezimm/npmfunctions/dist/Services/Users/IUserInterfaces';

import { IExStorageState, IItemDetail, IEXStorageList, IEXStorageBatch, IBatchData, IUserSummary, ITypeInfo, IDuplicateFile, IDuplicateInfo, IIconArray } from '../../IExStorageState';

import { IDataOptions, IUiOptions } from '../../IExStorageProps';

export interface IEsItemsProps {

      // 0 - Context
      // wpContext: WebPartContext;
      // pageContext: PageContext;

      // WebpartElement: HTMLElement;   //Size courtesy of https://www.netwoven.com/2018/11/13/resizing-of-spfx-react-web-parts-in-different-scenarios/

      // tenant: string;
  
      theSite: ISite;
      pickedWeb : IPickedWebBasic;
      pickedList? : IEXStorageList;
  
      // currentUser: IUser;
  
      items: IItemDetail[];
      sharedItems: IItemDetail[];
      itemsAreDups: boolean; //Set true if these items are "duplicates".  This will change the filename text to folder name because the filenames are all the same when it's a dup.
      itemsAreFolders: boolean; 
      childrenAreDups?: boolean; //Set true if these items are "duplicates".  This will change the filename text to folder name because the filenames are all the same when it's a dup.
        
      duplicateInfo?: IDuplicateInfo;

      icons: IIconArray[];

      heading: string;

      emptyItemsElements?: any[]; //Will pick any of these elements to randomly display
      // batches: IEXStorageBatch[];

      dataOptions: IDataOptions;
      uiOptions: IUiOptions;

      showHeading?: boolean;  //defaults to true

}
