
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

export interface IPerformanceSettings {
    fetchCount: number;
    fetchCountMobile: number;
    restFilter: string;
    minDataDownload: boolean;
}

export type IScaleMethod = 'slider' | 'blink' | 'pivot' | 'other' | 'na' | 'TBD';

export interface IGridchartsProps {

      // 0 - Context
      description: string;

      WebpartElement?: HTMLElement;   //Size courtesy of https://www.netwoven.com/2018/11/13/resizing-of-spfx-react-web-parts-in-different-scenarios/
      gridData?: ICSSChartSeries;
      color?: 'green' | 'red' | 'blue' | 'theme';

      pageContext: PageContext;
      wpContext: WebPartContext;
  
      allowOtherSites?: boolean; //default is local only.  Set to false to allow provisioning parts on other sites.
  
      allowRailsOff?: boolean;
      allowSettings?: boolean;
  
      tenant: string;
      urlVars: {};
      today: ITheTime;
  
      parentListWeb?: string;
      parentListURL?: string;
      parentListTitle?: string;
      listName : string;

      dateColumn: string;
      monthGap: string;

      valueColumn: string;
      searchColumns: string[];
      valueType: string;
      valueOperator: string;
      dropDownColumns: string[];
      metaColumns: string[];
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
  
      pivotSize: string;
      pivotFormat: string;
      pivotOptions: string;
      pivotTab: string;  //May not be needed because we have projectMasterPriority
  
      /**
       * 2020-09-08:  Add for dynamic data refiners.   onRefiner0Selected  -- callback to update main web part dynamic data props.
       */
      onRefiner0Selected?: any;
  
      gridStyles: {
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
      };
  
      //For DD
      handleSwitch: any;
      handleListPost: any;

      // 9 - Other web part options
      webPartScenario: string; //Choice used to create mutiple versions of the webpart. 
      showEarlyAccess: boolean;

}
