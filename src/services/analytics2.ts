import { sp } from '@pnp/sp';
import { Web, Items, } from '@pnp/sp/presets/all';

import { getHelpfullErrorV2, saveThisLogItem } from  '@mikezimm/npmfunctions/dist/Services/Logging/ErrorHandler';

import { getExpandColumns, getSelectColumns, IZBasicList, IPerformanceSettings, createFetchList, } from '@mikezimm/npmfunctions/dist/Lists/getFunctions';
import { sortObjectArrayByStringKey } from '@mikezimm/npmfunctions/dist/Services/Arrays/sorting';
import { IRailAnalytics } from '@mikezimm/npmfunctions/dist/Services/Arrays/grouping';
import { getFullUrlFromSlashSitesUrl } from '@mikezimm/npmfunctions/dist/Services/Strings/urlServices';
import { makeSmallTimeObject } from '@mikezimm/npmfunctions/dist/Services/Time/smallTimeObject';
import { msPerDay, msPerWk, msPerHr } from '@mikezimm/npmfunctions/dist/Services/Time/constants';

import { getBrowser, amIOnThisWeb, getWebUrlFromLink, getUrlVars,  } from '@mikezimm/npmfunctions/dist/Services/Logging/LogFunctions';
import { getCurrentPageLink, makeListLink, makeSiteLink, } from '@mikezimm/npmfunctions/dist/Services/Logging/LogFunctions';

import { saveAnalytics, ApplyTemplate_Rail_SaveTitle, ProvisionListsSaveTitle} from '@mikezimm/npmfunctions/dist/Services/Analytics/normal';

import { BaseErrorTrace } from './BaseErrorTrace';

/**
 * ILoadAnalytics can be created when the webpart loads the data so it's easy to pass
 */
export interface ILoadAnalytics {

  SiteID: string;  //Current site collection ID for easy filtering in large list
  WebID: string;  //Current web ID for easy filtering in large list
  SiteTitle: string; //Web Title
  ListID: string;  //Current list ID for easy filtering in large list
  ListTitle: string;
  
  TargetSite?: string;  //Saved as link column.  Displayed as Relative Url
  TargetList?: String;  //Saved as link column.  Displayed as Relative Url

}

/**
 * IZSentAnalytics can be created based on ILoadAnalytics when the webpart generates final data to save
 */
export interface IZSentAnalytics {

  loadProperties: ILoadAnalytics;

  Title: string;  //General Label used to identify what analytics you are saving:  such as Web Permissions or List Permissions.

  Result: string;  //Success or Error
  Setting?: string;  //Special settings

  zzzText1?: string; //Start-Now in some webparts
  zzzText2?: string; //Start-TheTime in some webparts
  zzzText3?: string; //Info1 in some webparts.  Simple category defining results.   Like Unique / Inherited / Collection
  zzzText4?: string; //Info2 in some webparts.  Phrase describing important details such as "Time to check old Permissions: 86 snaps / 353ms"
  zzzText5?: string;
  zzzText6?: string;
  zzzText7?: string;

  zzzNumber1?: number;
  zzzNumber2?: number;
  zzzNumber3?: number;
  zzzNumber4?: number;
  zzzNumber5?: number;
  zzzNumber6?: number;
  zzzNumber7?: number;

  zzzRichText1?: any;  //Used to store JSON objects for later use, will be stringified
  zzzRichText2?: any;
  zzzRichText3?: any;

  AnalyticsVersion?: string; //Not used in webparts, used in legacy html code
  CodeVersion?: string; //Not used in webparts, used in legacy html code

}

export interface ILink {
  Description: string;
  Url: string;
}

/**
 * This contains properties automatically added based on the current url
 */
export interface IZFullAnalytics extends IZSentAnalytics {

  loadProperties: any;  //To be removed in final object

  CollectionUrl?: string; // Should be target Site Collection Url

  PageURL: string;  //Url of page person is on
  getParams?: string;  //Parameters from url

  PageLink?: ILink;  // Saved as link column.  Displayed as Page Name
  SiteLink?: ILink;  //Saved as link column.  Displayed as full Url


  //These props were buried in loadProperties but get moved up to main object for saving.
  SiteID: string;  //Current site collection ID for easy filtering in large list
  WebID: string;  //Current web ID for easy filtering in large list
  SiteTitle: string; //Web Title
  ListID: string;  //Current list ID for easy filtering in large list
  ListTitle: string;
  
  TargetSite?: ILink;  //Saved as link column.  Displayed as Relative Url
  TargetList?: ILink;  //Saved as link column.  Displayed as Relative Url

  memory: string;
  browser: string;
  JSHeapSize: number;

  screen: string;  // Extra screen info in object link window.inner/outer sizes
  screenSize: string; // Basic dimensions 1080 x 1920
  device: string; 
  
}

// window.location properties
// host: "tenant.sharepoint.com"
// hostname: "tenant.sharepoint.com"
// href: "https://tenant.sharepoint.com/sites/WebPartDev/SitePages/ECStorage.aspx?debug=true&noredir=true&debugManifestsFile=https%3A%2F%2Flocalhost%3A4321%2Ftemp%2Fmanifests.js&allowOtherSites=true&scenario=dev"
// origin: "https://tenant.sharepoint.com"
// pathname: "/sites/WebPartDev/SitePages/ECStorage.aspx"
// protocol: "https:"
// search: "?debug=true&noredir=true&debugManifestsFile=https%3A%2F%2Flocalhost%3A4321%2Ftemp%2Fmanifests.js&allowOtherSites=true&scenario=dev"

export function getSiteCollectionUrlFromLink( link: string ) {
  if ( !link || link.length === 0 ) {
    link = window.location.pathname ;
  } else if ( link.indexOf('http') === 0 ) {
    link = link.replace( window.location.origin, '');
  }
  //At this point, link should be relative url /sites/collection....
  let parts = link.split('/');
  let collectionUrl = `${window.location.origin}/sites/${parts[2]}`;
  return collectionUrl;
}

export function getSizeLabel ( size: number) {
  return size > 1e9 ? `${ (size / 1e9).toFixed(1) } GB` : size > 1e6 ? `${ (size / 1e6).toFixed(1) } MB` : `${ ( size / 1e3).toFixed(1) } KB`;
}

export function saveAnalytics2 ( analyticsWeb: string, analyticsList: string, saveObject: IZSentAnalytics,  ) {
  let saveOjbectCopy: any = JSON.parse(JSON.stringify( saveObject )) ;
  let finalSaveObject: IZFullAnalytics = saveOjbectCopy;

  finalSaveObject.AnalyticsVersion = 'saveAnalytics2';

  delete finalSaveObject[ 'loadProperties' ];

  finalSaveObject.SiteID = saveObject.loadProperties.SiteID;  //Current site collection ID for easy filtering in large list
  finalSaveObject.WebID = saveObject.loadProperties.WebID;  //Current web ID for easy filtering in large list
  finalSaveObject.SiteTitle = saveObject.loadProperties.SiteTitle; //Web Title
  finalSaveObject.ListID = saveObject.loadProperties.ListID;  //Current list ID for easy filtering in large list
  finalSaveObject.ListTitle = saveObject.loadProperties.ListTitle;

  if ( typeof saveObject.zzzRichText1 === 'object' ) { 
    finalSaveObject.zzzRichText1 = JSON.stringify( saveObject.zzzRichText1 ); 
    console.log('Length of zzzRichText1:', finalSaveObject.zzzRichText1.length );
  } else if ( typeof saveObject.zzzRichText1 === 'string' ) { 
    finalSaveObject.zzzRichText1 = saveObject.zzzRichText1 ; 
  }

  if ( typeof saveObject.zzzRichText2 === 'object' ) { 
    finalSaveObject.zzzRichText2 = JSON.stringify( saveObject.zzzRichText2 ); 
    console.log('Length of zzzRichText2:', finalSaveObject.zzzRichText2.length );
  } else if ( typeof saveObject.zzzRichText2 === 'string' ) { 
    finalSaveObject.zzzRichText2 = saveObject.zzzRichText2 ; 
  }

  if ( typeof saveObject.zzzRichText3 === 'object' ) { 
    finalSaveObject.zzzRichText3 = JSON.stringify( saveObject.zzzRichText3 ); 
    console.log('Length of zzzRichText3:', finalSaveObject.zzzRichText3.length );
  } else if ( typeof saveObject.zzzRichText3 === 'string' ) { 
    finalSaveObject.zzzRichText3 = saveObject.zzzRichText3 ; 
  }

  //Convert TargetSite to actual link object
  if ( typeof saveObject.loadProperties.TargetSite === 'string' ) {
    finalSaveObject.TargetSite = {
      'Url': saveObject.loadProperties.TargetSite,
      'Description': saveObject.loadProperties.SiteTitle,
    };
  }

  //Convert TargetSite to actual link TargetList
  if ( typeof saveObject.loadProperties.TargetList === 'string' ) {
    finalSaveObject.TargetList = {
      'Url': saveObject.loadProperties.TargetList,
      'Description': saveObject.loadProperties.ListTitle,
    };
  }

  //Create CollectionUrl string from TargetSite
  if ( saveObject.loadProperties.TargetSite && !finalSaveObject.CollectionUrl ) {

    if ( typeof saveObject.loadProperties.TargetSite === 'string' ) {
      finalSaveObject.CollectionUrl = getSiteCollectionUrlFromLink(saveObject.loadProperties.TargetSite); // Should be target Site Collection Url

    } else if ( typeof saveObject.loadProperties.TargetSite === 'object' ) {
      finalSaveObject.CollectionUrl = getSiteCollectionUrlFromLink(saveObject.loadProperties.TargetSite); // Should be target Site Collection Url

    }
  }

  //Add current Page Link and Url
  finalSaveObject.PageLink = getCurrentPageLink();
  finalSaveObject.PageURL = finalSaveObject.PageLink.Url;

  //Add parameters
  finalSaveObject.getParams = getUrlVars().join(' & ');

  let SiteLink = getWebUrlFromLink( '' , 'abs');
  let SiteTitle = SiteLink.substring(SiteLink.lastIndexOf("/") + 1);

  finalSaveObject.SiteLink = {
    'Url': SiteLink,
    'Description': SiteTitle,
  };

  /**
   * Get Memory usage information
   */
  //Courtesy of https://trackjs.com/blog/monitoring-javascript-memory/
  let memoryObj = window.performance['memory'];

  if ( memoryObj ) {
    memoryObj.usedPerTotal = memoryObj.totalJSHeapSize && memoryObj.totalJSHeapSize !== 0 ? memoryObj.usedJSHeapSize / memoryObj.totalJSHeapSize : null;
    memoryObj.totalPerLimit = memoryObj.jsHeapSizeLimit && memoryObj.jsHeapSizeLimit !== 0 ? memoryObj.totalJSHeapSize / memoryObj.jsHeapSizeLimit : null;
    memoryObj.usedPerLimit = memoryObj.jsHeapSizeLimit && memoryObj.jsHeapSizeLimit !== 0 ? memoryObj.usedJSHeapSize / memoryObj.jsHeapSizeLimit : null;
    memoryObj.Limit = getSizeLabel( memoryObj.jsHeapSizeLimit );
    memoryObj.Total = getSizeLabel( memoryObj.totalJSHeapSize );
    memoryObj.Used = getSizeLabel( memoryObj.usedJSHeapSize );

    finalSaveObject.memory = JSON.stringify( memoryObj );
    finalSaveObject.browser = 'Chromium';
    finalSaveObject.JSHeapSize = memoryObj.totalJSHeapSize;

  } else {
    finalSaveObject.browser = 'Not Chromium';
  }

  /**
   * Get screen information
   */

  let screen = null;
  if ( window && window.screen ) {
    screen = {
      screenWidth: window.screen.width,
      screenHeight: window.screen.height,

      outerWidth: window.outerWidth,
      outerHeight: window.outerHeight,

      innerWidth: window.innerWidth,
      innerHeight: window.innerHeight,

      ratio: window.screen.width / window.screen.height,
      aspect: getAspectRatio( window.screen.width, window.screen.height ),

    };
  }

  finalSaveObject.screen = screen;
  finalSaveObject.screenSize = `${innerHeight} x ${innerWidth}`;

  /**
   * get device information
   */
  
  let OSName = null;
  if ( navigator && navigator.appVersion ) {
    if ( navigator.appVersion.indexOf("Win")!=-1) OSName="Windows";
    else if (navigator.appVersion.indexOf("Mac")!=-1) OSName="MacOS";
    else if (navigator.appVersion.indexOf("X11")!=-1) OSName="UNIX";
    else if (navigator.appVersion.indexOf("Linux")!=-1) OSName="Linux";
  }

  let device = {
    OSName: OSName,
  };

  finalSaveObject.device = JSON.stringify( device );

  saveThisLogItem( analyticsWeb, analyticsList, finalSaveObject );

}

export function getAspectRatio( width: number, height: number ) {
  if ( height === 0 || width === 0 ) {
    return 'na';
  } else {
    let result = `${width} / ${ height }`;
    let ratio = roundRatio( width / height );
    if ( ratio === roundRatio(16/9 ) ) { result = '16 / 9' ; }
    else if ( ratio === roundRatio(9/16 ) ) { result = '9 / 16' ; }
    else if ( ratio === roundRatio(4/3 ) ) { result = '4 / 3' ; }
    else if ( ratio === roundRatio(3/4 ) ) { result = '3 / 4' ; }
    else if ( ratio === roundRatio(21/9 ) ) { result = '21 / 9' ; }
    else if ( ratio === roundRatio(9/21 ) ) { result = '9 / 21' ; }
    else if ( ratio === roundRatio(14/9 ) ) { result = '14 / 9' ; }
    else if ( ratio === roundRatio(9/14 ) ) { result = '9 / 14' ; }
    else if ( ratio === roundRatio(18/9 ) ) { result = '18 / 9' ; }
    else if ( ratio === roundRatio(9/18 ) ) { result = '9 / 18' ; }
    else if ( ratio === roundRatio(23/16 ) ) { result = '23 / 16' ; } // Ipad Air 4
    else if ( ratio === roundRatio(16/23 ) ) { result = '16 / 23' ; } // Ipad Air 4
    else if ( ratio === roundRatio(19.5/9 ) ) { result = '19.5 / 9' ; } // Iphone 11-12-XR
    else if ( ratio === roundRatio(9/23 ) ) { result = '9 / 19.5' ; } // Iphone 11-12-XR
    else if ( ratio === roundRatio(4/5 ) ) { result = '4 / 5' ; }
    else if ( ratio === roundRatio(5/4 ) ) { result = '5 / 4' ; }
    else if ( ratio === roundRatio(32/9 ) ) { result = '32 / 9' ; }
    else if ( ratio === roundRatio(9/32 ) ) { result = '9 / 32' ; }
    return result;
  }
}

export function roundRatio( num: number ) {
  if ( num < 1 ) {
    return round3decimals( num );
  } else {
    return round1decimals( num );
  }
}


export function round3decimals( num: number ) {
  return Math.round(num * 1000) / 1000;
}
export function round2decimals( num: number ) {
  return Math.round(num * 100) / 100;
}
export function round1decimals( num: number ) {
  return Math.round(num * 10) / 10;
}
