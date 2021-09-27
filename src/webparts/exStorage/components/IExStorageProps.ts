
import { sp, Views, IViews, ISite } from "@pnp/sp/presets/all";

import { WebPartContext } from '@microsoft/sp-webpart-base';

import { PageContext } from '@microsoft/sp-page-context';

import { escape } from '@microsoft/sp-lodash-subset';

import { IPickedWebBasic, IPickedList, }  from '@mikezimm/npmfunctions/dist/Lists/IListInterfaces';

import { IUser } from '@mikezimm/npmfunctions/dist/Services/Users/IUserInterfaces';
import { IChoiceGroupOptionStyleProps } from "office-ui-fabric-react";

import { IWebpartBannerProps, IWebpartBannerState } from './HelpInfo/banner/bannerProps';

export interface IDataOptions {
  useMediaTags: boolean;

} 

export interface IUiOptions {

  
  showListDropdown: boolean;

  showSystemLists: boolean;

  excludeListTitles: string;

  /** quickCloseItem:
   * Set to true to easily pop open and close item panel.
   * Set to false to force you do do proper dismiss (click X or outside panel)
   */

  quickCloseItem: boolean;

  /**
   * 400 is default
   * Setting higher will allow you to see more items on the screen but it becomes sluggish during search
   */
  maxVisibleItems: number;
   
} 

export interface IExStorageProps {

      // 0 - Context
      wpContext: WebPartContext;
      pageContext: PageContext;

      WebpartElement: HTMLElement;   //Size courtesy of https://www.netwoven.com/2018/11/13/resizing-of-spfx-react-web-parts-in-different-scenarios/

        // 1 - Analytics options
      useListAnalytics: boolean;
      analyticsWeb: string;
      analyticsList: string;
      tenant: string;
      urlVars: {};

      bannerProps: IWebpartBannerProps;

      parentWeb: string;
      listTitle: string;
  
      allowOtherSites?: boolean; //default is local only.  Set to false to allow provisioning parts on other sites.
      pickedWeb : IPickedWebBasic;
      theSite: ISite;

      isLoaded: boolean;
  
      currentUser: IUser;
  
      WebpartHeight: number;
      WebpartWidth: number;

      dataOptions: IDataOptions;
      uiOptions: IUiOptions;

}
