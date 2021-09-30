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

import { BaseErrorTrace } from './BaseErrorTrace';

/**
 * Be sure to update your analyticsList and analyticsWeb in en-us.js strings file
 * @param theProps 
 * @param theState 
 */

export function saveListory (analyticsWeb, analyticsList, SiteLink, webTitle, saveTitle, TargetSite, TargetList, itemInfo1, itemInfo2, result, listory, Setting, fields, views ) {

    //Do nothing if either of these strings is blank
    if (!analyticsList) { return ; }
    if (!analyticsWeb) { return ; }

    
    let saveItem: any ={
        Title: saveTitle,
        PageLink: getCurrentPageLink(),
        ListTitle: itemInfo1,
    };

    let startTime = getTheCurrentTime();
    saveItem.zzzText1 = startTime.now;
    saveItem.zzzText2 = startTime.theTime;

    saveItem.getParams = getUrlVars().join(' & ');
    saveItem.Setting = Setting;

    saveItem.zzzRichText1 = listory ? JSON.stringify(listory) : null;
    saveItem.zzzRichText2 = fields ? JSON.stringify(fields) : null;
    saveItem.zzzRichText3 = views ? JSON.stringify(views) : null;

    let tempSite = TargetSite.split('|');
    TargetSite = tempSite[0];
    saveItem.WebID = tempSite[1] ? tempSite[1] : null;
    saveItem.CollectionUrl = tempSite[2] ? tempSite[2] : null;
    saveItem.SiteID = tempSite[3] ? tempSite[3] : null;
    saveItem.zzzText5 = saveItem.SiteID;

    SiteLink = getWebUrlFromLink( SiteLink, 'abs' );

    if ( webTitle === '' || !webTitle ) {
        saveItem.SiteTitle = SiteLink.substring(SiteLink.lastIndexOf("/") + 1);
    }

    saveItem.SiteLink = {
        'Url': SiteLink && SiteLink.indexOf('http') === 0 ? SiteLink : window.location.origin + SiteLink,
        'Description': saveItem.SiteTitle ,
    };
    
    saveItem.TargetSite = makeSiteLink( TargetSite, saveItem.SiteTitle );

    saveItem.TargetList = makeListLink( TargetList, webTitle );

    saveItem.PageLink = getCurrentPageLink();

    saveThisLogItem( analyticsWeb, analyticsList, saveItem );

}

// /**
//  * Be sure to update your analyticsList and analyticsWeb in en-us.js strings file

//  */
// export const ApplyTemplate_Rail_SaveTitle = 'Apply Template Rail';
// export const ApplyTemplate_Page_SaveTitle = 'Apply Template Page';
// export const ProvisionListsSaveTitle = 'Provision Lists';
// export function saveAnalytics (analyticsWeb, analyticsList, SiteLink, webTitle, saveTitle, TargetSite, TargetList, itemInfo1, itemInfo2, result, RichTextJSON1, Setting, RichTextJSON2, RichTextJSON3 ) {

//     //Do nothing if either of these strings is blank
//     if (!analyticsList) { return ; }
//     if (!analyticsWeb) { return ; }

//     let saveItem: any ={
//         Title: saveTitle,
//         Result: result,
//         PageLink: getCurrentPageLink(),
//     };

//     let startTime = getTheCurrentTime();
//     saveItem.zzzText1 = startTime.now;
//     saveItem.zzzText2 = startTime.theTime;

//     let TargetListValues = TargetList ? TargetList.split('|') : [null];
    
//     saveItem.getParams = getUrlVars().join(' & ');
//     saveItem.Setting = Setting;

//     // console.log('saveAnalytics StringifyActionJson: ', RichTextJSON1, RichTextJSON2, RichTextJSON3 );
//     saveItem.zzzRichText1 = RichTextJSON1 ? JSON.stringify(RichTextJSON1) : null;
//     saveItem.zzzRichText2 = RichTextJSON2 ? JSON.stringify(RichTextJSON2) : null;
//     saveItem.zzzRichText3 = RichTextJSON3 ? JSON.stringify(RichTextJSON3) : null;

//     if ( analyticsList === strings.analyticsListRailsGroups || 
//         analyticsList === strings.analyticsListRailsApply ||
//         analyticsList === strings.analyticsListPermissionsHistory 
//         ) { //Rails Off
//         saveItem.ListTitle = itemInfo1;

//         let infos2 = itemInfo2 ? itemInfo2.split('|') : [ ];

//         saveItem.zzzText3 = infos2[0];

//         saveItem.zzzText7 = infos2[1] ? parseInt(infos2[1]) < 10 ? '0' + infos2[1] : infos2[1] : null ; //stepOrder

//         saveItem.zzzNumber4 = infos2[2] ? parseInt( infos2[2] ) : null ;
//         saveItem.zzzNumber5 = infos2[3] ? parseInt( infos2[3] ) : null ;

//         saveItem.zzzText1 = infos2[4] ? infos2[4] : null ;
//         saveItem.zzzText4 = infos2[5] ? infos2[5] : null;

//         let tempSite = TargetSite ? TargetSite.split('|') : [];
//         TargetSite = tempSite[0] ? tempSite[0] : null;
//         saveItem.WebID = tempSite[1] ? tempSite[1] : null;
//         saveItem.CollectionUrl = tempSite[2] ? tempSite[2] : null;
//         saveItem.SiteID = tempSite[3] ? tempSite[3] : null;
//         saveItem.zzzText5 = saveItem.SiteID;

//         //Add List ID if it's available
//         if ( TargetListValues.length > 0 && TargetListValues[1] ) { saveItem.ListID = TargetListValues[1] ; }

//         let tempTitle = saveTitle.split('|');
//         saveItem.zzzText6 = tempTitle[1] ? tempTitle[1] : null;//Get scope - site or list

//     } else {
//         saveItem.zzzText3 = itemInfo1;
//         saveItem.zzzText4 = itemInfo2;

//     }

//     SiteLink = getWebUrlFromLink( SiteLink , 'abs');

//     if ( webTitle === '' || !webTitle ) {
//         saveItem.SiteTitle = SiteLink.substring(SiteLink.lastIndexOf("/") + 1);
//     }

//     saveItem.SiteLink = {
//         'Url': SiteLink && SiteLink.indexOf('http') === 0 ? SiteLink : window.location.origin + SiteLink,
//         'Description': saveItem.SiteTitle ,
//     };
    
//     saveItem.TargetSite = TargetSite ? makeSiteLink( TargetSite, saveItem.SiteTitle ) : null ;

//     saveItem.TargetList = TargetList ? makeListLink( TargetListValues[0], webTitle ) : null;

//     saveThisLogItem( analyticsWeb + '', analyticsList + '', saveItem );

// }

// /**
//  * 
//  * @param analyticsWeb 
//  * @param analyticsList 
//  * @param WebID 
//  * @param ListID 
//  * @param fetchOnlyThisList :  Set to true in order to add the list ID to the rest filter to return only relavent items
//  */
// export async function fetchAnalytics( analyticsWeb: string, analyticsList: string, WebID: string, ListID: string = null, fetchOnlyThisList: boolean = false, theseColumns: any[] = [], top: number = 5000 ) {
//     //Do nothing if either of these strings is blank
//     if (!analyticsList) { return ; }
//     if (!analyticsWeb) { return ; }

//     let items: IRailAnalytics[] = [];

//     let allColumns : any = theseColumns.length > 0 ? JSON.parse(JSON.stringify( theseColumns )) :
//     [ 'Created','Modified','Author/Name','Author/Id','Author/Title','Id',
//         'Title', 'zzzRichText1', 'zzzRichText2', 'zzzRichText3', 'getParams',
//         'zzzNumber1', 'zzzNumber2', 'zzzNumber3', 'zzzNumber4', 'zzzNumber5',
//         'zzzText1', 'zzzText2', 'zzzText3', 'zzzText4', 'zzzText5', 'zzzText6', 'zzzText7',
//         'PageLink', 'SiteLink', 'SiteTitle', 'TargetSite', 'Result',
//         'TargetList', 'ListTitle', 'Setting','WebID','SiteID','CollectionUrl', 'ListID'
//     ];

//     let expColumns : any = getExpandColumns(allColumns);

//     analyticsWeb = getFullUrlFromSlashSitesUrl( analyticsWeb );

//     try {
//         let web = Web(analyticsWeb);
//         let restFilter = "WebID eq '" + WebID + "'";
        
//         if ( fetchOnlyThisList === true && ListID && ListID.length > 0 ) {
//             restFilter += " and ListID eq '" + ListID + "'";
//         }
        
//         items = await web.lists.getByTitle(analyticsList).items.select(allColumns).expand(expColumns).filter( restFilter ).top(top).orderBy('Id',false).get();

//     } catch (e) {
//         console.log('e',getHelpfullErrorV2(e, true,true, [ BaseErrorTrace , 'Failed', 'Fetch Analytics', ].join('|') ) );

//     }

//     return items ;

// }


// /**
//  * This function is for automatically saving permissions from a web, list or library to list for later comparison.
//  * In Easy Contents, it's fired upon viewing rail function to view list permissions.
//  * It's also intended to be used in Pivot Tiles when clicking to view list and web permissions.
//  * 
//  * It does require the list and web with the correct struture to save and then be recoverd in this webpart for comparison.
//  * 
//  * So it's only going to execute in certain tenanats.
//  * If you see this and want to re-purpose it, update the function to suit your needs and adjust the window.location.origin check
//  * 
//  * Best practice is just to update your site and list Url in strings:
//  *  Or just create the site:  SharePointAssist
//  *  And create the list:  Assists
//  *  And add the columns listed below in the save item
//     "analyticsListPermissionsHistory": "PermissionsHistory",
//  * 
// */

// export async function savePermissionHistory ( analyticsWeb, analyticsList, SiteLink, webTitle, saveTitle, TargetSite, TargetList, itemInfo1, itemInfo2, result, RichTextJSON1, Setting, RichTextJSON2, RichTextJSON3, userName ) {

//     let prefetchStart = new Date();
//     let pickedWebguid = TargetSite.split('|')[1];
//     let theListId = TargetList.split('|')[1];
//     let fetchColumns = [ 'Created','Modified','Author/Name','Author/Id','Author/Title','Id',
//         'Title','zzzRichText3', 'Result', 'WebID','SiteID','CollectionUrl', 'ListID'
//     ];

//     /**
//      * In this case I chose to fetch the last 200 items to compare and see if the current user had any history.
//      * If not, it will save the permissions even if it has not changed since another user checked.
//      * Choose 200 to make sure that it should almost always be enough to check the last day's worth of items.
//      */
//     let items: IRailAnalytics[] = await fetchAnalytics( analyticsWeb, analyticsList , pickedWebguid, theListId, true, fetchColumns, 200 );

//     let lastIsSame: any = false;
//     // console.log('RichTextJSON3', RichTextJSON3 );
//     RichTextJSON3 = JSON.stringify(RichTextJSON3);
//     console.log('RichTextJSON3 length = ', RichTextJSON3.length, RichTextJSON3 );

//     let lastTimeCurrentUserSaved = null;
//     let saveThisSnapshot: any = false;
//     let checkedOtherUsers = false;
//     let foundMyItem = false;

//     items.map( ( item, index)  => {
//         if ( saveThisSnapshot === false && foundMyItem === false ) { 
//             let itemFromCurrentUser = false;
//             let itemAny: any = item;
//             if ( itemAny.Author.Name === userName ) {
//                 console.log('You saved this item:', item );
//                 lastTimeCurrentUserSaved = new Date(item.Created);
//                 itemFromCurrentUser = true;
//                 foundMyItem = true;
//             }

//             let userDeltaTime = itemFromCurrentUser === false ? null : prefetchStart.getTime() - lastTimeCurrentUserSaved.getTime() ;

//             if ( itemFromCurrentUser === true && userDeltaTime > msPerDay ) {  //one day = 24*60*60*1000
//                 saveThisSnapshot = true;  //Save this item if the current user has not saved permissions in last 24 hours

//             } else if ( checkedOtherUsers === false && lastIsSame === false ) { //this check happens if 
//                 //This section checks if the current item has the same permissions settings as the current check
//                 let zzzRichText3 = item.zzzRichText3.replace(/\\\"/g,'"');
//                 zzzRichText3 = zzzRichText3.slice( 1, -1 );//Have to remove the leading and trailing "" 
//                 console.log('zzzRichText3 length = ', zzzRichText3.length, zzzRichText3 );
//                 if ( zzzRichText3.length === RichTextJSON3.length && zzzRichText3 === RichTextJSON3 ) { 
//                     lastIsSame = true;
//                 }
//                 checkedOtherUsers = true;  //This is used so it only does this check one time.
//             }
//         }
//     });

//     //Final check for other conditions to save.
//     if ( saveThisSnapshot === true ){
//         //No need for further checks

//     } if ( foundMyItem === false ) { //automatically save since this user never saved permissions
//         saveThisSnapshot = true;

//     } else if ( lastIsSame === false ) { //Save if the last item is not the same as current permissions
//         saveThisSnapshot = true;

//     }

//     let prefetchEnd = new Date();
//     let timeToPreFetch = prefetchEnd.getTime() - prefetchStart.getTime();
//     itemInfo2 += '||Time to check old Permissions: ' + items.length + ' snaps'  + ' / ' + timeToPreFetch + 'ms' ;

//     console.log('savePermissionHistory lastIsSame', lastIsSame );

//     if ( saveThisSnapshot === true ) {
//         RichTextJSON1 = JSON.stringify(RichTextJSON1);
//         RichTextJSON2 = JSON.stringify(RichTextJSON2);
    
//         saveAnalytics( analyticsWeb, analyticsList , //analyticsWeb, analyticsList,
//             SiteLink, webTitle,//serverRelativeUrl, webTitle,
//             saveTitle, TargetSite, TargetList, //saveTitle, TargetSite, TargetList
//             itemInfo1, itemInfo2, result, //itemInfo1, itemInfo2, result, 
//             RichTextJSON1, Setting, RichTextJSON2, RichTextJSON3 ); //richText, Setting, richText2, richText3
//     }

// }


/**
 * This function is for automatically creating a item in our Teams' request list in SharePoint.
 * Initially it's fired upon completing rail functions to auto-document support incidents.
 * 
 * So it's only going to execute in certain tenanats.
 * If you see this and want to re-purpose it, update the function to suit your needs and adjust the window.location.origin check
 * 
 * Best practice is just to update your site and list Url in strings:
 *  Or just create the site:  SharePointAssist
 *  And create the list:  Assists
 *  And add the columns listed below in the save item
    "requestListSite": "/sites/SharePointAssist",
    "requestListList": "Assists",
 * 
*/

export function saveAssist ( analyticsWeb, analyticsList, SiteLink, webTitle, saveTitle, TargetSite, TargetList, itemInfo1, itemInfo2: string[], result, RichTextJSON1, Setting, RichTextJSON2, RichTextJSON3 ) {

    if ( window.location.origin.indexOf( 'utoliv.sharepoint.com') < 0 && window.location.origin.indexOf( 'clickster.sharepoint')  < 0 ) { return ; }

    if (!analyticsList) { return ; }
    if (!analyticsWeb) { return ; }

    SiteLink = getWebUrlFromLink( SiteLink, 'abs' );

    let location = makeListLink( TargetList, webTitle );

    // let startTime = makeSmallTimeObject( null );
    // let localTimeString = startTime.theTime;
    let localTimeString = new Date();
    let StatusComments = RichTextJSON1 ? typeof RichTextJSON1 === 'string' ? RichTextJSON1 : JSON.stringify(RichTextJSON1).replace('\"','') : null;
    let ScopeArray: string[] = itemInfo2;
    let saveItem: any ={
        Title: saveTitle,
        Scope:  { results: ScopeArray },  //Need to add scope back in as multi-select choice.
        Status: '4. Completed', //Choice
        Complexity: '0 Automation', //Choice
        StatusComments: StatusComments, //Mulit-Line Text (plain text)
        StartDate: localTimeString, //Date-Time
        EndDate: localTimeString, //Date-Time
        TargetCompleteDate: localTimeString, //Date-Time
        Location: location, //Link
    };

    saveThisLogItem( analyticsWeb + '', analyticsList + '', saveItem );

}

export function saveAnalyticsX (theTime) {

    let analyticsList = "TilesCycleTesting";
    let currentTime = theTime;
    let web = Web('https://mcclickster.sharepoint.com/sites/Templates/SiteAudit/');

    web.lists.getByTitle(analyticsList).items.add({
        'Title': 'Pivot-Tiles x1asdf',
        'zzzText1': currentTime.now,      
        'zzzText2': currentTime.theTime,
        'zzzNumber1': currentTime.milliseconds,

        }).then((response) => {
        //Reload the page
            //location.reload();
        }).catch((e) => {
        //Throw Error
            alert(e);
    });


}

export function saveTheTime () {
    let theTime = getTheCurrentTime();
    saveAnalyticsX(theTime);

    return theTime;

}

export function getTheCurrentTime () {

    const now = new Date();
    const theTime = now.getHours() + ":" + now.getMinutes() + ":" + now.getSeconds() + "." + now.getMilliseconds();
    let result : any = {
        'now': now,
        'theTime' : theTime,
        'milliseconds' : now.getMilliseconds(),
    };

    return result;

}
