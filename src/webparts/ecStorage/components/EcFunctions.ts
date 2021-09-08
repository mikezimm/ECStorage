
import { Web, SiteGroups, SiteGroup, ISiteGroups, ISiteGroup, ISiteGroupInfo, IPrincipalInfo, PrincipalType, PrincipalSource } from "@pnp/sp/presets/all";

import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-groups/web";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { PageContext } from '@microsoft/sp-page-context';
import { mergeAriaAttributeValues, IconNames } from "office-ui-fabric-react";


/***
 *    d888888b .88b  d88. d8888b.  .d88b.  d8888b. d888888b      d8b   db d8888b. .88b  d88.      d88888b db    db d8b   db  .o88b. d888888b d888888b  .d88b.  d8b   db .d8888. 
 *      `88'   88'YbdP`88 88  `8D .8P  Y8. 88  `8D `~~88~~'      888o  88 88  `8D 88'YbdP`88      88'     88    88 888o  88 d8P  Y8 `~~88~~'   `88'   .8P  Y8. 888o  88 88'  YP 
 *       88    88  88  88 88oodD' 88    88 88oobY'    88         88V8o 88 88oodD' 88  88  88      88ooo   88    88 88V8o 88 8P         88       88    88    88 88V8o 88 `8bo.   
 *       88    88  88  88 88~~~   88    88 88`8b      88         88 V8o88 88~~~   88  88  88      88~~~   88    88 88 V8o88 8b         88       88    88    88 88 V8o88   `Y8b. 
 *      .88.   88  88  88 88      `8b  d8' 88 `88.    88         88  V888 88      88  88  88      88      88b  d88 88  V888 Y8b  d8    88      .88.   `8b  d8' 88  V888 db   8D 
 *    Y888888P YP  YP  YP 88       `Y88P'  88   YD    YP         VP   V8P 88      YP  YP  YP      YP      ~Y8888P' VP   V8P  `Y88P'    YP    Y888888P  `Y88P'  VP   V8P `8888Y' 
 *                                                                                                                                                                              
 *                                                                                                                                                                              
 */

import { IUser } from '@mikezimm/npmfunctions/dist/Services/Users/IUserInterfaces';
import { doesObjectExistInArrayInt, } from '@mikezimm/npmfunctions/dist/Services/Arrays/checks';
import { sortObjectArrayByNumberKey } from '@mikezimm/npmfunctions/dist/Services/Arrays/sorting';
import { getSiteAdmins } from '@mikezimm/npmfunctions/dist/Services/Users/userServices';   //groupUsers = await getSiteAdmins( webURL, false);
import { getHelpfullErrorV2 } from '@mikezimm/npmfunctions/dist/Services/Logging/ErrorHandler';

import { getPrincipalTypeString } from '@mikezimm/npmfunctions/dist/Services/Users/userServices';
import { getFullUrlFromSlashSitesUrl } from '@mikezimm/npmfunctions/dist/Services/Strings/urlServices';
import { IEcStorageState, IECStorageList, IECStorageBatch } from './IEcStorageState';
import { escape } from '@microsoft/sp-lodash-subset';

import { IPickedWebBasic, IPickedList, }  from '@mikezimm/npmfunctions/dist/Lists/IListInterfaces';

 /**
  * These properties throw error on fetching.
  * ,"ServerRedirectedPreviewURL", "SharedWithInternal"
  */
 const thisSelect = ['ID','FileRef','FileLeafRef','Author/Title','Editor/Title','Modified','Created','CheckoutUserId','HasUniqueRoleAssignments','Title','FileSystemObjectType','FileSizeDisplay','FileLeafRef','LinkFilename'];
 const thisExpand = ['Author','Editor'];

 export async function getStorageItems( pickedWeb: IPickedWebBasic , pickedList: IECStorageList, addTheseItemsToState: any, setProgress: any, ) {

  let webURL = pickedWeb.url;
  let listTitle = pickedList.Title;
  let batchSize = 500;

  let items: any = null;

  let isLoaded = false;

  let errMessage = '';
  let thisWebInstance = null;

  let batches: IECStorageBatch[] = [];
 
  try {
    thisWebInstance = Web(webURL);
    let thisListObject = thisWebInstance.lists.getByTitle( listTitle );
    setProgress( 0 , pickedList.ItemCount, 'Getting ' + 'first' + ' batches of items' );
    try {

      let fetchStart = new Date();
      let startMs = fetchStart.getTime();
      items = await thisListObject.items.select(thisSelect).expand(thisExpand).top(batchSize).filter('').getPaged(); 

      batches = batches.concat( createThisBatch( items, startMs ) );
      for ( let i = 1; i < 3 ; i++ ) {
        if ( items.hasNext ) {
          let thisBatchStart = i * batchSize ;
          setProgress( thisBatchStart , pickedList.ItemCount, `Fetching ${thisBatchStart} of ${ pickedList.ItemCount } items` );
          fetchStart = new Date();
          startMs = fetchStart.getTime();
          items = await items.getNext();
          batches = batches.concat( createThisBatch( items, startMs ) );
        }
      }
      


    } catch( e ) {
      let helpfulErrorEnd = [ webURL, listTitle, null, null ].join('|');
      errMessage = getHelpfullErrorV2(e, false, true, [ 'BaseErrorTrace' , 'Failed', 'GetStorage ~ 59', helpfulErrorEnd ].join('|') );
    }

  } catch (e) {
    let helpfulErrorEnd = [ webURL, listTitle, null, null ].join('|');
    errMessage = getHelpfullErrorV2(e, false, true, [ 'BaseErrorTrace' , 'Failed', 'GetStorage ~ 64', helpfulErrorEnd ].join('|') );
 
  }


  console.log('getStorageItems:', batches );
  addTheseItemsToState( batches );

  return { batches };
 
 }

 function createThisBatch ( items: any, start: number ) {
        
    let fetchEnd = new Date();
    let endMs = fetchEnd.getTime();
    let duration = endMs - start;
    let count = items && items.results ? items.results.length : 0;

    let batch: IECStorageBatch = {
      start: start,
      end: endMs,
      duration: duration,
      msPerItem: count > 0 ? duration / count : 0,
      count: count,
      errMessage: '',
      id: '',
      items: [].concat( items.results ),
      hasNext: items.hasNext,
    };

    return batch;

 }
 export function analyzeStorage( oldItems: any[] ) {
  let items: any[] = [];

  return oldItems;

 }