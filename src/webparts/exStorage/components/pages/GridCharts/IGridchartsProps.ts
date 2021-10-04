
/**
 * 
 * 
 * Official Community Imports
 * 
 * 
 */

import { PageContext } from '@microsoft/sp-page-context';
import { WebPartContext } from '@microsoft/sp-webpart-base';


/**
 * 
 * 
 * @mikezimm/npmfunctions/dist/ Imports
 * 
 * 
 */

import { ITheTime} from '@mikezimm/npmfunctions/dist/Services/Time/Interfaces';

import { ICSSChartSeries,  } from '@mikezimm/npmfunctions/dist/CSSCharts/ICSSCharts';

import { IPickedWebBasic, IPickedList, }  from '@mikezimm/npmfunctions/dist/Lists/IListInterfaces';

import { IEXStorageList } from '../../IExStorageState';

import { IDataOptions, IUiOptions } from '../../IExStorageProps';

/**
 * 
 * 
 * Services Imports
 * 
 * 
 */


 
/**
 * 
 * 
 * Helper Imports
 * 
 * 
 */


/**
 * 
 * This Component Imports
 * 
 * 
 */
export interface IGridStyles {

    cellColor: string;
    yearStyles: string;
    monthStyles: string;
    dayStyles: string;
    cellStyles: string;
    cellhoverInfoColor: string;
    other: string;
    
    squareCustom: string;
    squareColor: string;
    emptyColor: string;
    backGroundColor: string;
    monthGap: string; 

}


export type IValueOperator = 'Sum' | 'Min' | 'Max' | 'Avg' | 'Count';

export interface IGridColumns {
  dateColumn: string;
  valueColumn: string;
  searchColumns: string[];
  valueType: string;
  valueOperators: IValueOperator[];
  dropDownColumns: string[];
  metaColumns: string[];
}

export interface IPerformanceSettings {
    fetchCount: number;
    fetchCountMobile: number;
    restFilter: string;
    minDataDownload: boolean;
}

export type IScaleMethod = 'slider' | 'blink' | 'pivot' | 'other' | 'na' | 'TBD';


export interface IGridchartsProps {

      WebpartElement?: HTMLElement;   //Size courtesy of https://www.netwoven.com/2018/11/13/resizing-of-spfx-react-web-parts-in-different-scenarios/

      pageContext: PageContext;
      wpContext: WebPartContext;
  
      allowOtherSites?: boolean; //default is local only.  Set to false to allow provisioning parts on other sites.
  
      allowRailsOff?: boolean;
      allowSettings?: boolean;
  
      tenant: string;
      urlVars: {};
      today: ITheTime;

      items: any[]; //Added to bring in any items into this component
  
      parentListWeb?: string;
      parentListURL?: string;

      //VVV Added from Extreme Storage for Items
      pickedWeb? : IPickedWebBasic;
      pickedList? : IEXStorageList;

      dataOptions: IDataOptions;
      uiOptions: IUiOptions;

      //^^^ Added from Extreme Storage for Items 

      esItemsHeading: string;

      parentListTitle?: string;
      listName : string;

      columns: IGridColumns;

      enableSearch: boolean;

      allLoaded: boolean;
  
      scaleMethod: IScaleMethod;

      performance: IPerformanceSettings;
  
      parentListFieldTitles: string;
  
      // 1 - Analytics options
      useListAnalytics: boolean;
      analyticsWeb?: string;
      analyticsList?: string;
  
      /**    
       * 'parseBySemiColons' |
       * 'groupBy10s' |  'groupBy100s' |  'groupBy1000s' |  'groupByMillions' |
       * 'groupByDays' |  'groupByMonths' |  'groupByYears' |
       * 'groupByUsers' | 
       * 
       * rules string formatted as JSON : [ string[] ]  =  [['parseBySemiColons''groupByMonths'],['groupByMonths'],['groupByUsers']]
       * [ ['parseBySemiColons''groupByMonths'],
       * ['groupByMonths'],
       * ['groupByUsers'] ]
       * 
      */
  
      // 6 - User Feedback:
      //progress: IMyProgress;
  
      WebpartHeight?:  number;    //Size courtesy of https://www.netwoven.com/2018/11/13/resizing-of-spfx-react-web-parts-in-different-scenarios/
      WebpartWidth?:   number;    //Size courtesy of https://www.netwoven.com/2018/11/13/resizing-of-spfx-react-web-parts-in-different-scenarios/
  
      /**
       * 2020-09-08:  Add for dynamic data refiners.   onRefiner0Selected  -- callback to update main web part dynamic data props.
       */
      onRefiner0Selected?: any;
  
      gridStyles: IGridStyles;

      // 9 - Other web part options
      webPartScenario: string; //Choice used to create mutiple versions of the webpart. 
      // showEarlyAccess: boolean;
      refreshId?: string; //used to trigger redraw of grid

}
