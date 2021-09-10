
import { Web, SiteGroups, SiteGroup, ISiteGroups, ISiteGroup, ISiteGroupInfo, IPrincipalInfo, PrincipalType, PrincipalSource } from "@pnp/sp/presets/all";

import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-groups/web";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { PageContext } from '@microsoft/sp-page-context';
import { mergeAriaAttributeValues, IconNames } from "office-ui-fabric-react";

import "@pnp/sp/search";
import { Search, Suggest } from "@pnp/sp/search";

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
import { IEcStorageState, IECStorageList, IECStorageBatch, IItemDetail, IBatchData, ILargeFiles, IUserFiles, IOldFiles, IUserSummary, IFileType } from './IEcStorageState';
import { escape } from '@microsoft/sp-lodash-subset';

import { IPickedWebBasic, IPickedList, }  from '@mikezimm/npmfunctions/dist/Lists/IListInterfaces';

 /**
  * These properties throw error on fetching.
  * ,"ServerRedirectedPreviewURL", "SharedWithInternal"
  */

/**
 * These size fields throw error when fetching:
 * 'File_x0020_Size','SMTotalSize','File_x0020_Size','SMTotalFileStreamSize', 'DocumentSummarySize','tp_UIVersion','_UIVersionString','odata.UIVersionString'
 * 
  * FileSystemObjectType:  https://docs.microsoft.com/en-us/previous-versions/office/sharepoint-server/ee537053(v=office.15)#members
  *  File=0; Folder=1; Web=0
 */
 const thisSelect = ['*','ID','FileRef','FileLeafRef','Author/Title','Editor/Title','Author/Name','Editor/Name','Modified','Created','CheckoutUserId','HasUniqueRoleAssignments','Title','FileSystemObjectType','FileSizeDisplay','FileLeafRef','LinkFilename','OData__UIVersion','OData__UIVersionString','DocIcon'];
 const thisExpand = ['Author','Editor'];
  export const batchSize = 500;

export function createLargeFiles() :ILargeFiles {
  return {  
    GT10G: [],
    GT01G: [],
    GT100M: [],
    GT10M: [],
  };
}

export function createOldFiles () :IOldFiles {
  return {  
    Age5Yr: [],
    Age4Yr: [],
    Age3Yr: [],
    Age2Yr: [],
    Age1Yr: [],
  };
}

export function createUserFiles (): IUserFiles {
  return {  
    large: createLargeFiles(),
    oldCreated: createOldFiles(),
    oldModified: createOldFiles(),
    items: [],
  };
}

export function createThisUser( detail : IItemDetail, userId: number, userTitle: string ) :IUserSummary {

  let userSummary: IUserSummary = {
    userId: userId,
    userTitle: userTitle,
    userFirst: null,
    userLast: null,
    createCount: 0,
    modifyCount: 0,
    folderCreateCount: 0,
    createTotalSize: 0,
    modifyTotalSize: 0,
    createTotalSizeGB: 0,
    modifyTotalSizeGB: 0,
    createSizes: [],
    modifiedSizes: [],
  };

  return userSummary;

}

export function updateThisEditor ( detail : IItemDetail, userSummary: IUserSummary ) {

  if ( userSummary.userId === detail.editorId ) {
    if ( detail.isFolder === true ) {
      //do nothing
    } else {
      userSummary.modifyCount ++;
      userSummary.modifyTotalSize += detail.size;
      userSummary.modifiedSizes.push( detail.size );

    }
  }
  return userSummary;

}

export function updateThisAuthor ( detail : IItemDetail, userSummary: IUserSummary ) {

  if ( userSummary.userId === detail.authorId ) {
    if ( detail.isFolder === true ) {
      userSummary.folderCreateCount ++;

    } else {
      userSummary.createCount ++;
      userSummary.createTotalSize += detail.size;
      userSummary.createSizes.push( detail.size );

    }
  }

  return userSummary;

}

export function createThisType ( docIcon: string ) :IFileType {

  let thisType: IFileType = {
    type: docIcon,
    count: 0,
    size: 0,
    sizeGB: 0,
    items: [],
    sizes: [],
    createdMs: [],
    modifiedMs: [],
  };

  return thisType;

}

export function updateThisType ( thisType: IFileType, detail : IItemDetail, ) : IFileType {

  thisType.count ++;
  thisType.size += detail.size;

  thisType.items.push( detail );
  thisType.sizes.push(detail.size);

  thisType.createdMs.push( detail.createMs ) ;
  thisType.modifiedMs.push( detail.modMs ) ;

  return thisType;

}

//IBatchData, ILargeFiles, IUserFiles, IOldFiles
export function createBatchData ():IBatchData {
  return {  
    count: 0,
    size: 0,
    sizeGB: 0,
    typeList: [],
    types: [],
    duplicateNames: [],
    duplicates: [],
    large: createLargeFiles(),
    oldCreated: createOldFiles(),
    oldModified: createOldFiles(),
    currentUser: createUserFiles(),
    folders: [],
    creatorIds: [],
    editorIds: [],
    allUsersIds: [],
    allUsers: [],
    uniqueRolls: [],
  };
}

 export async function getStorageItems( pickedWeb: IPickedWebBasic , pickedList: IECStorageList, fetchCount: number, userId: number, addTheseItemsToState: any, setProgress: any, ) {

  userId = 6;  //REMOVE THIS LINE>>> USED FOR TESTING ONLY

  let webURL = pickedWeb.url;
  let listTitle = pickedList.Title;

  let items: any = null;

  let isLoaded = false;

  let errMessage = '';
  let thisWebInstance = null;

  let batches: IECStorageBatch[] = [];
 
  if ( fetchCount > 0 ) {
    try {
    
      // set the url for search
      // const searcher = Search(webURL);
  
      // This testing did not return anything I can understand that looks like a result.
      // this can accept any of the query types (text, ISearchQuery, or SearchQueryBuilder)
      // const results = await searcher("Frauenhofer");
      // console.log('Test searcher results', results);
  
      thisWebInstance = Web(webURL);
      let thisListObject = thisWebInstance.lists.getByTitle( listTitle );
      setProgress( 0 , pickedList.ItemCount, 'Getting ' + 'first' + ' batches of items' );
      try {
  
        let fetchStart = new Date();
        let startMs = fetchStart.getTime();
        items = await thisListObject.items.select(thisSelect).expand(thisExpand).top(batchSize).filter('').getPaged(); 
  
        batches = batches.concat( createThisBatch( items, startMs, 0 ) );
        for ( let i = 1; i < 150 ; i++ ) {
          let thisBatchStart = i * batchSize ;
          if ( items.hasNext && fetchCount > thisBatchStart ) {
            setProgress( thisBatchStart , fetchCount, `Fetching ${thisBatchStart} of ${ fetchCount } items` );
            fetchStart = new Date();
            startMs = fetchStart.getTime();
            items = await items.getNext();
            batches = batches.concat( createThisBatch( items, startMs, i ) );
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
  }

  let batchData = createBatchData();

  let currentDate = new Date();
  let currentYear = currentDate.getFullYear();

  let analyzeStart = new Date();
  let startMs2 = analyzeStart.getTime();

  //These variables are used to develop ceilings for metrics
  let largest: IItemDetail = null;
  let oldestCreate: IItemDetail = null;
  let oldestModified: IItemDetail = null;
  let userLargest: IItemDetail = null;
  let userOldestCreate: IItemDetail = null;
  let userOldestModified: IItemDetail = null;

  batches.map( batch=> {
    batch.items.map( ( item, itemIndex )=> {

      //Get item summary
      let detail: IItemDetail = createGenericItemDetail( batch.index , itemIndex, item, userId );

      batchData.count ++;
      batchData.size += detail.size;

      //Build up Type list
      let typeIndex = batchData.typeList.indexOf( detail.docIcon );

      if ( typeIndex < 0 ) {
        batchData.typeList.push( detail.docIcon );
        typeIndex = batchData.typeList.length - 1;
        batchData.types.push( createThisType(detail.docIcon) );
      }
      batchData.types[ typeIndex ] = updateThisType( batchData.types[ typeIndex ], detail );

      //Build up Duplicate list



      //Get index of authorId in array of all authorIds
      let createUserIndex = batchData.creatorIds.indexOf( detail.authorId );
      if ( createUserIndex === -1 ) { 
        batchData.creatorIds.push( detail.authorId  );
        createUserIndex = batchData.creatorIds.length -1;
      }

      //Get index of editor in array of all editorIds
      let editUserIndex = batchData.editorIds.indexOf( detail.editorId  );
      if ( editUserIndex === -1 ) { 
        batchData.editorIds.push( detail.editorId  );
        editUserIndex = batchData.editorIds.length -1;
      }

      //Get index of author in array of all allIds - to get the allUser Item for later use
      let createUserAllIndex = batchData.allUsersIds.indexOf( detail.authorId );
      if ( createUserAllIndex === -1 ) { 
        batchData.allUsersIds.push( detail.authorId  );
        batchData.allUsers.push( createThisUser( detail, detail.authorId, detail.authorTitle )  );
        createUserAllIndex = batchData.allUsers.length -1;
      }

      //Get index of editor in array of all allIds - to get the allUser Item for later use
      let editUserAllIndex = batchData.allUsersIds.indexOf( detail.editorId  );
      if ( editUserAllIndex === -1 ) { 
        batchData.allUsersIds.push( detail.editorId  );
        batchData.allUsers.push( createThisUser( detail, detail.editorId, detail.editorTitle )  );
        editUserAllIndex = batchData.allUsers.length -1;
      }

      batchData.allUsers[ createUserAllIndex ] = updateThisAuthor( detail, batchData.allUsers[ createUserAllIndex ]);
      batchData.allUsers[ editUserAllIndex ] = updateThisEditor( detail, batchData.allUsers[ editUserAllIndex ]);

      //Set default high items
      if ( !detail.isFolder ) {
        if ( largest === null ) {  largest = detail ; } else if ( detail.size > largest.size ) { largest = detail ; }
        if ( oldestCreate === null ) {  oldestCreate = detail ; } else if ( detail.createMs < oldestCreate.createMs ) { oldestCreate = detail ; }
        if ( oldestModified === null ) {  oldestModified = detail ; } else if ( detail.modMs < oldestModified.modMs ) { oldestModified = detail ; }
        if ( detail.currentUser === true ) {  
          //Set user high items
          if ( userLargest === null ) { 
            userLargest = detail ;
            userOldestCreate = detail ;
            userOldestModified = detail ;
          }
          if ( detail.size > userLargest.size ) { userLargest = detail ; }
          if ( detail.createMs < userOldestCreate.createMs ) { userOldestCreate = detail ; }
          if ( detail.modMs < userOldestModified.modMs ) { userOldestModified = detail ; }
        }
      }

      if ( detail.currentUser === true ) { batchData.currentUser.items.push ( detail ) ; } 
      if ( detail.isFolder === true ) { batchData.folders.push ( detail ) ; } 
      if ( detail.uniquePerms === true ) { batchData.uniqueRolls.push ( detail ) ; } 

      //Add to large bucket
      if ( detail.size > 1e10 ) { 
        batchData.large.GT10G.push ( detail ) ;
        if ( detail.currentUser === true ) { batchData.currentUser.large.GT10G.push ( detail ) ; }   

       } else if ( detail.size > 1e9 ) { 
        batchData.large.GT01G.push ( detail ) ; 
        if ( detail.currentUser === true ) { batchData.currentUser.large.GT01G.push ( detail ) ; }  

      } else if ( detail.size > 1e8 ) { 
        batchData.large.GT100M.push ( detail ) ; 
        if ( detail.currentUser === true ) { batchData.currentUser.large.GT100M.push ( detail ) ; }   

      } else if ( detail.size > 1e7 ) { 
        batchData.large.GT10M.push ( detail ) ; 
        if ( detail.currentUser === true ) { batchData.currentUser.large.GT10M.push ( detail ) ; }    

      }

      if ( detail.createYr < currentYear - 4 ) { 
        batchData.oldCreated.Age5Yr.push ( detail ) ;
        if ( detail.currentUser === true ) { batchData.currentUser.oldCreated.Age5Yr.push ( detail ) ; }    
       }
      else if ( detail.createYr < currentYear - 3 ) { 
        batchData.oldCreated.Age4Yr.push ( detail ) ; 
        if ( detail.currentUser === true ) { batchData.currentUser.oldCreated.Age4Yr.push ( detail ) ; }  
      }
      else if ( detail.createYr < currentYear - 2 ) { 
        batchData.oldCreated.Age3Yr.push ( detail ) ; 
        if ( detail.currentUser === true ) { batchData.currentUser.oldCreated.Age3Yr.push ( detail ) ; }  
      }
      else if ( detail.createYr < currentYear - 1 ) { 
        batchData.oldCreated.Age2Yr.push ( detail ) ; 
        if ( detail.currentUser === true ) { batchData.currentUser.oldCreated.Age2Yr.push ( detail ) ; }  
      }
      else if ( detail.createYr < currentYear - 0 ) { 
        batchData.oldCreated.Age1Yr.push ( detail ) ; 
        if ( detail.currentUser === true ) { batchData.currentUser.oldCreated.Age1Yr.push ( detail ) ; }  
      }

      if ( detail.modYr < currentYear - 4 ) { 
        batchData.oldModified.Age5Yr.push ( detail ) ;
        if ( detail.currentUser === true ) { batchData.currentUser.oldModified.Age5Yr.push ( detail ) ; }    
       }
      else if ( detail.modYr < currentYear - 3 ) { 
        batchData.oldModified.Age4Yr.push ( detail ) ; 
        if ( detail.currentUser === true ) { batchData.currentUser.oldModified.Age4Yr.push ( detail ) ; }  
      }
      else if ( detail.modYr < currentYear - 2 ) { 
        batchData.oldModified.Age3Yr.push ( detail ) ; 
        if ( detail.currentUser === true ) { batchData.currentUser.oldModified.Age3Yr.push ( detail ) ; }  
      }
      else if ( detail.modYr < currentYear - 1 ) { 
        batchData.oldModified.Age2Yr.push ( detail ) ; 
        if ( detail.currentUser === true ) { batchData.currentUser.oldModified.Age2Yr.push ( detail ) ; }  
      }
      else if ( detail.modYr < currentYear - 0 ) { 
        batchData.oldModified.Age1Yr.push ( detail ) ; 
        if ( detail.currentUser === true ) { batchData.currentUser.oldModified.Age1Yr.push ( detail ) ; }  
      }

    });
  });

  batchData.sizeGB += ( batchData.size / 1e9 );
  batchData.types.map( docType => {
    docType.sizeGB = docType.size/1e9;
  });

  batchData.allUsers.map( user => {
    user.createTotalSizeGB = user.createTotalSize / 1e9;
    user.modifyTotalSizeGB = user.modifyTotalSize / 1e9;
  });

  let analyzeEnd = new Date();
  let endMs2 = analyzeEnd.getTime();
  let analyzeMs = endMs2 - startMs2;

  let fetchMs = 0;
  let totalLength = 0;
  batches.map ( batch => { 
    fetchMs += batch.duration;
    totalLength += batch.items.length;
  });

  let batchInfo = {
    batches: batches,
    batchData: batchData,
    fetchMs: fetchMs,
    analyzeMs: analyzeMs,
    totalLength: totalLength,
  };

  console.log('getStorageItems: fetchMs', fetchMs );
  console.log('getStorageItems: analyzeMs', analyzeMs );
  console.log('getStorageItems: totalLength', totalLength );

  console.log('getStorageItems: batches', batches );
  console.log('getStorageItems: batchData', batchData );

  addTheseItemsToState( batchInfo );

  return { batches };
 
 }

 function createGenericItemDetail ( batchIndex:  number, itemIndex:  number, item: any, userId ) : IItemDetail {
  let created = new Date(item.Created);
  let modified = new Date(item.Modified);

  let createYr = created.getFullYear();
  let modYr = modified.getFullYear();

  let currentUser = item.AuthorId === userId ? true : false;
  currentUser = item.EditorId === userId ? true : currentUser;

  let itemDetail: IItemDetail = {
    batch: batchIndex, //index of the batch in state.batches
    index: itemIndex, //index of item in state.batches[batch].items
    value: null, //value to highlight/sort for this detail
    created: created,
    modified: modified,
    authorId: item.AuthorId,
    editorId: item.EditorId,
    authorTitle: item.Author.Title,
    editorTitle: item.Editor.Title,
    FileLeafRef: item.FileLeafRef,
    FileRef: item.FileRef,
    id: item.Id,
    currentUser: currentUser,
    size: item.FileSizeDisplay ? parseInt(item.FileSizeDisplay) : 0,
    sizeMB: item.FileSizeDisplay ? Math.round( parseInt(item.FileSizeDisplay) / 1e6 * 100) / 100 : 0,
    createYr: createYr,
    modYr: modYr,
    bucket: `${createYr}-${modYr}`,
    createMs: created.getTime(),
    modMs: modified.getTime(),
  };


  if ( item.CheckoutUserId ) { itemDetail.checkedOutId = item.CheckoutUserId; }
  if ( item.HasUniqueRoleAssignments ) { itemDetail.uniquePerms = item.HasUniqueRoleAssignments; }
  if ( item.FileSystemObjectType === 1 ) { itemDetail.isFolder = true; }

  if ( item.DocIcon ) { 
    itemDetail.docIcon = item.DocIcon; 
  } else if ( itemDetail.isFolder === true ) {
    itemDetail.docIcon = 'folder'; 
  }

  return itemDetail;

 }

 function createThisBatch ( results: any, start: number, batchIndex: number ) {
        
    let fetchEnd = new Date();
    let endMs = fetchEnd.getTime();
    let duration = endMs - start;
    let items = results.results;
    let count = items && items.length > 0 ? items.length : 0;
    let firstCreated = items && items.length > -1 ? new Date( items[0].Created ) : null;
    let lastCreated = items && items.length > -1 ? new Date( items[items.length - 1 ].Created ) : null;

    let batch: IECStorageBatch = {
      index: batchIndex,
      start: start,
      end: endMs,
      duration: duration,
      msPerItem: count > 0 ? duration / count : 0,
      firstCreated: firstCreated,
      lastCreated: lastCreated,
      count: count,
      errMessage: '',
      id: '',
      items: [].concat( items ),
      hasNext: results.hasNext,
    };

    return batch;

 }

 export function analyzeStorage( oldItems: any[] ) {
  let items: any[] = [];

  return oldItems;

 }