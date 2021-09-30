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

  if ( typeof saveObject.zzzRichText1 === 'object' ) { finalSaveObject.zzzRichText1 = JSON.stringify( saveObject.zzzRichText1 ); }
  if ( typeof saveObject.zzzRichText2 === 'object' ) { finalSaveObject.zzzRichText1 = JSON.stringify( saveObject.zzzRichText2 ); }
  if ( typeof saveObject.zzzRichText3 === 'object' ) { finalSaveObject.zzzRichText1 = JSON.stringify( saveObject.zzzRichText3 ); }

  finalSaveObject.zzzRichText1 = saveObject.zzzRichText1 ? JSON.stringify(saveObject.zzzRichText1) : null;
  finalSaveObject.zzzRichText2 = saveObject.zzzRichText2 ? JSON.stringify(saveObject.zzzRichText2) : null;
  finalSaveObject.zzzRichText3 = saveObject.zzzRichText3 ? JSON.stringify(saveObject.zzzRichText3) : null;

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
    finalSaveObject.browser = 'Chrome';
    finalSaveObject.JSHeapSize = memoryObj.totalJSHeapSize;

  } else {
    finalSaveObject.browser = 'Not Chrome';
  }

  saveThisLogItem( analyticsWeb, analyticsList, finalSaveObject );

}