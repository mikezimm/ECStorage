
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
import { sortNumberArray, sortObjectArrayByChildNumberKey } from '@mikezimm/npmfunctions/dist/Services/Arrays/sorting';
import { expandArray } from '@mikezimm/npmfunctions/dist/Services/Arrays/manipulation';
// import { sortObjectArrayByChildNumberKey, sortNumberArray } from '@mikezimm/npmfunctions/dist/Services/Arrays/sorting';


import { getSiteAdmins } from '@mikezimm/npmfunctions/dist/Services/Users/userServices';   //groupUsers = await getSiteAdmins( webURL, false);
import { getHelpfullErrorV2 } from '@mikezimm/npmfunctions/dist/Services/Logging/ErrorHandler';

import { getPrincipalTypeString } from '@mikezimm/npmfunctions/dist/Services/Users/userServices';
import { getFullUrlFromSlashSitesUrl } from '@mikezimm/npmfunctions/dist/Services/Strings/urlServices';

import { IPickedWebBasic, IPickedList, }  from '@mikezimm/npmfunctions/dist/Lists/IListInterfaces';
import { msPerDay, msPerWk }  from '@mikezimm/npmfunctions/dist/Services/Time/constants';

import { updateNextOpenIndex } from '@mikezimm/npmfunctions/dist/Services/Arrays/manipulation';
import { getSizeLabel, getCountLabel } from '@mikezimm/npmfunctions/dist/Services/Math/basicOperations'; 

import { IExStorageState, IEXStorageList, IEXStorageBatch, IItemDetail, IBatchData, ILargeFiles, IOldFiles, IUserSummary, IFileType, 
    IDuplicateFile, IBucketSummary, IUserInfo, ITypeInfo, IFolderInfo, IDuplicateInfo, IFolderDetail, IAllItemTypes, IBucketType } from './IExStorageState';

import { IDataOptions, IUiOptions } from './IExStorageProps';

import { escape } from '@microsoft/sp-lodash-subset';


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

 /***
 *     .o88b.  .d88b.  d8b   db .d8888. d888888b 
 *    d8P  Y8 .8P  Y8. 888o  88 88'  YP `~~88~~' 
 *    8P      88    88 88V8o 88 `8bo.      88    
 *    8b      88    88 88 V8o88   `Y8b.    88    
 *    Y8b  d8 `8b  d8' 88  V888 db   8D    88    
 *     `Y88P'  `Y88P'  VP   V8P `8888Y'    YP    
 *                                               
 *                                               
 */

import { sharedWithSelect, sharedWithExpand, processSharedItems } from './Sharing/SharingFunctions2';
import { IItemSharingInfo, ISharingEvent, ISharedWithUser } from './Sharing/ISharingInterface';


const domainEmail = window.location.hostname.replace('.sharepoint','');

 const thisSelect = ['*','ID','FileRef','FileLeafRef','Author/Title','Editor/Title','Author/Name','Editor/Name','Modified','Created','CheckoutUserId','HasUniqueRoleAssignments','Title','FileSystemObjectType','FileSizeDisplay','FileLeafRef','LinkFilename','OData__UIVersion','OData__UIVersionString','DocIcon'];

 //Preservation Hold Library errors out if you try to select the Title.  All other properties work.
 const presHoldSelect = ['*','ID','FileRef','FileLeafRef','Author/Title','Editor/Title','Author/Name','Editor/Name','Modified','Created','CheckoutUserId','HasUniqueRoleAssignments','FileSystemObjectType','FileSizeDisplay','FileLeafRef','LinkFilename','OData__UIVersion','OData__UIVersionString','DocIcon'];


 const thisExpand = ['Author','Editor'];
  export const batchSize = 500;

 /***
 *     d888b  d88888b d888888b       .o88b. db    db d8888b. d8888b. d88888b d8b   db d888888b      db    db d88888b  .d8b.  d8888b. 
 *    88' Y8b 88'     `~~88~~'      d8P  Y8 88    88 88  `8D 88  `8D 88'     888o  88 `~~88~~'      `8b  d8' 88'     d8' `8b 88  `8D 
 *    88      88ooooo    88         8P      88    88 88oobY' 88oobY' 88ooooo 88V8o 88    88          `8bd8'  88ooooo 88ooo88 88oobY' 
 *    88  ooo 88~~~~~    88         8b      88    88 88`8b   88`8b   88~~~~~ 88 V8o88    88            88    88~~~~~ 88~~~88 88`8b   
 *    88. ~8~ 88.        88         Y8b  d8 88b  d88 88 `88. 88 `88. 88.     88  V888    88            88    88.     88   88 88 `88. 
 *     Y888P  Y88888P    YP          `Y88P' ~Y8888P' 88   YD 88   YD Y88888P VP   V8P    YP            YP    Y88888P YP   YP 88   YD 
 *                                                                                                                                   
 *                                                                                                                                   
 */

  function getCurrentYear(){
    let currentDate = new Date();
    let currentYear = currentDate.getFullYear();
    return currentYear;
  }

/***
 *     .o88b. d8888b. d88888b  .d8b.  d888888b d88888b      .d8888. db    db .88b  d88. .88b  d88.  .d8b.  d8888b. db    db 
 *    d8P  Y8 88  `8D 88'     d8' `8b `~~88~~' 88'          88'  YP 88    88 88'YbdP`88 88'YbdP`88 d8' `8b 88  `8D `8b  d8' 
 *    8P      88oobY' 88ooooo 88ooo88    88    88ooooo      `8bo.   88    88 88  88  88 88  88  88 88ooo88 88oobY'  `8bd8'  
 *    8b      88`8b   88~~~~~ 88~~~88    88    88~~~~~        `Y8b. 88    88 88  88  88 88  88  88 88~~~88 88`8b      88    
 *    Y8b  d8 88 `88. 88.     88   88    88    88.          db   8D 88b  d88 88  88  88 88  88  88 88   88 88 `88.    88    
 *     `Y88P' 88   YD Y88888P YP   YP    YP    Y88888P      `8888Y' ~Y8888P' YP  YP  YP YP  YP  YP YP   YP 88   YD    YP    
 *                                                                                                                          
 *                                                                                                                          
 */

  export function createBucketSummary( title: string, bucket: IBucketType ): IBucketSummary {
    let summary: IBucketSummary = {
      title: title,
      bucket: bucket,
      count: 0,
      size: 0,
      sizeGB: 0,
      sizeToCountRatio: 0,
      sizeLabel: '',
      countP: 0,
      sizeP: 0,
      userIds: [],
      userTitles: [],
      ranges: {
        firstCreateMs: 1e20,
        lastCreateMs:  0,
        firstModifiedMs:  1e20,
        lastModifiedMs:  0,
        createRange: ``,
        modifyRange: ``,
        firstAllMs: 0,
        lastAllMs: 0,
        rangeAll: ``,
      }
    };
    return summary;

  }

  /***
 *    db    db d8888b. d8888b.  .d8b.  d888888b d88888b      .d8888. db    db .88b  d88. .88b  d88.  .d8b.  d8888b. db    db 
 *    88    88 88  `8D 88  `8D d8' `8b `~~88~~' 88'          88'  YP 88    88 88'YbdP`88 88'YbdP`88 d8' `8b 88  `8D `8b  d8' 
 *    88    88 88oodD' 88   88 88ooo88    88    88ooooo      `8bo.   88    88 88  88  88 88  88  88 88ooo88 88oobY'  `8bd8'  
 *    88    88 88~~~   88   88 88~~~88    88    88~~~~~        `Y8b. 88    88 88  88  88 88  88  88 88~~~88 88`8b      88    
 *    88b  d88 88      88  .8D 88   88    88    88.          db   8D 88b  d88 88  88  88 88  88  88 88   88 88 `88.    88    
 *    ~Y8888P' 88      Y8888D' YP   YP    YP    Y88888P      `8888Y' ~Y8888P' YP  YP  YP YP  YP  YP YP   YP 88   YD    YP    
 *                                                                                                                           
 *                                                                                                                           
 */

  export function updateBucketSummary( summary: IBucketSummary, detail: IItemDetail ): IBucketSummary {
    summary.count ++;
    summary.size += detail.size;
    summary.sizeGB = summary.size/1e9;
    summary.sizeLabel = getSizeLabel( summary.size ) ;
    if ( summary.userIds.indexOf( detail.authorId ) < 0 ) { summary.userIds.push( detail.authorId ) ; }
    if ( summary.userIds.indexOf( detail.editorId ) < 0 ) { summary.userIds.push( detail.editorId ) ; }
    if ( summary.userTitles.indexOf( detail.authorTitle ) < 0 ) { summary.userTitles.push( detail.authorTitle ) ; }
    if ( summary.userTitles.indexOf( detail.editorTitle ) < 0 ) { summary.userTitles.push( detail.editorTitle ) ; }

    let rangeChanged = false;
    // debugger;
    if ( detail.createMs < summary.ranges.firstCreateMs ) { summary.ranges.firstCreateMs = detail.createMs ; rangeChanged = true ; }
    if ( detail.createMs > summary.ranges.lastCreateMs ) { summary.ranges.lastCreateMs = detail.createMs ; rangeChanged = true ; }
    if ( detail.modMs < summary.ranges.firstModifiedMs ) { summary.ranges.firstModifiedMs = detail.modMs ; rangeChanged = true ; }
    if ( detail.modMs > summary.ranges.lastModifiedMs ) { summary.ranges.lastModifiedMs = detail.modMs ; rangeChanged = true ; }
    // console.log('BucketSummary:', rangeChanged, detail.id, detail.createMs, detail.modMs );
    if ( rangeChanged === true ) {
      let firstCreateMs = new Date(summary.ranges.firstCreateMs);
      let lastCreateMs = new Date(summary.ranges.lastCreateMs);
      let firstModifiedMs = new Date(summary.ranges.firstModifiedMs);
      let lastModifiedMs = new Date(summary.ranges.lastModifiedMs);

      let firstCreateLocal = firstCreateMs.toLocaleDateString();
      let lastCreateLocal = lastCreateMs.toLocaleDateString();
      let firstModifiedLocal = firstModifiedMs.toLocaleDateString();
      let lastModifiedLocal = lastModifiedMs.toLocaleDateString();

      summary.ranges.createRange = firstCreateLocal !== lastCreateLocal ? `${firstCreateLocal} - ${lastCreateLocal}` : firstCreateLocal;
      summary.ranges.modifyRange = firstModifiedLocal !== lastModifiedLocal ? `${firstModifiedLocal} - ${lastModifiedLocal}` : firstModifiedLocal;

      summary.ranges.firstAllMs = Math.min(...[summary.ranges.firstCreateMs, summary.ranges.lastCreateMs, summary.ranges.firstModifiedMs, summary.ranges.lastModifiedMs]);
      summary.ranges.lastAllMs = Math.max(...[summary.ranges.firstCreateMs, summary.ranges.lastCreateMs, summary.ranges.firstModifiedMs, summary.ranges.lastModifiedMs]);

      let firstAllMs = new Date(summary.ranges.firstAllMs);
      let lastAllMs = new Date(summary.ranges.lastAllMs);

      let firstAllLocal = firstAllMs.toLocaleDateString();
      let lastAllLocal = lastAllMs.toLocaleDateString();

      summary.ranges.rangeAll = firstAllLocal !== lastAllLocal ? `${firstAllLocal} - ${lastAllLocal}` : firstAllLocal ;

    }
    return summary;

  }

  export function updateBucketSummaryPercents( summary: IBucketSummary, compare: IBucketSummary ): IBucketSummary {
    summary.sizeGB = summary.size / 1e9 ;
    summary.sizeP = 100 * summary.size / compare.size ;
    summary.countP = 100 * summary.count / compare.count;
    summary.sizeLabel = getSizeLabel( summary.size ) ;
    summary.sizeToCountRatio = summary.sizeP / summary.countP;
    return summary;

  }


  /***
 *     .o88b. d8888b. d88888b  .d8b.  d888888b d88888b      db       .d8b.  d8888b.  d888b  d88888b 
 *    d8P  Y8 88  `8D 88'     d8' `8b `~~88~~' 88'          88      d8' `8b 88  `8D 88' Y8b 88'     
 *    8P      88oobY' 88ooooo 88ooo88    88    88ooooo      88      88ooo88 88oobY' 88      88ooooo 
 *    8b      88`8b   88~~~~~ 88~~~88    88    88~~~~~      88      88~~~88 88`8b   88  ooo 88~~~~~ 
 *    Y8b  d8 88 `88. 88.     88   88    88    88.          88booo. 88   88 88 `88. 88. ~8~ 88.     
 *     `Y88P' 88   YD Y88888P YP   YP    YP    Y88888P      Y88888P YP   YP 88   YD  Y888P  Y88888P 
 *                                                                                                  
 *                                                                                                  
 */

export function createLargeFiles() :ILargeFiles {
  return {  
    GT10G: [],
    GT01G: [],
    GT100M: [],
    GT10M: [],
    summary: createBucketSummary( `Files BIGGER than 100MB`, 'Large Files' ),
  };
}

/***
 *     .o88b. d8888b. d88888b  .d8b.  d888888b d88888b       .d88b.  db      d8888b. 
 *    d8P  Y8 88  `8D 88'     d8' `8b `~~88~~' 88'          .8P  Y8. 88      88  `8D 
 *    8P      88oobY' 88ooooo 88ooo88    88    88ooooo      88    88 88      88   88 
 *    8b      88`8b   88~~~~~ 88~~~88    88    88~~~~~      88    88 88      88   88 
 *    Y8b  d8 88 `88. 88.     88   88    88    88.          `8b  d8' 88booo. 88  .8D 
 *     `Y88P' 88   YD Y88888P YP   YP    YP    Y88888P       `Y88P'  Y88888P Y8888D' 
 *                                                                                   
 *                                                                                   
 */
export function createOldFiles () :IOldFiles {
  return {  
    Age5Yr: [],
    Age4Yr: [],
    Age3Yr: [],
    Age2Yr: [],
    Age1Yr: [],
    summary: createBucketSummary( `Files created before ${( getCurrentYear() - 1 )}`, 'Old Files' ),
  };
}

/***
 *     .o88b. d8888b. d88888b  .d8b.  d888888b d88888b      db    db .d8888. d88888b d8888b. 
 *    d8P  Y8 88  `8D 88'     d8' `8b `~~88~~' 88'          88    88 88'  YP 88'     88  `8D 
 *    8P      88oobY' 88ooooo 88ooo88    88    88ooooo      88    88 `8bo.   88ooooo 88oobY' 
 *    8b      88`8b   88~~~~~ 88~~~88    88    88~~~~~      88    88   `Y8b. 88~~~~~ 88`8b   
 *    Y8b  d8 88 `88. 88.     88   88    88    88.          88b  d88 db   8D 88.     88 `88. 
 *     `Y88P' 88   YD Y88888P YP   YP    YP    Y88888P      ~Y8888P' `8888Y' Y88888P 88   YD 
 *                                                                                           
 *                                                                                           
 */

export function createThisUser( detail : IItemDetail, userId: number, userTitle: string, userName: string ) :IUserSummary {
  
  let sharedNameSplit = userName ? userName.split('|') : [];
  let sharedName = sharedNameSplit.length > 1 ? sharedNameSplit[ 2 ].replace( domainEmail, '' ) : userName;

  let userSummary: IUserSummary = {
    userId: userId,
    userTitle: userTitle,
    sharedName: sharedName,
    userFirst: null,
    userLast: null,

    createCount: 0,
    createSizes: [],
    createTotalSize: 0,
    createTotalSizeLabel: '',
    createTotalSizeGB: 0,
    createSizeRank: 0,
    createCountRank: 0,
    oldCreated: createOldFiles(),

    modifyCount: 0,
    modifyTotalSize: 0,
    modifyTotalSizeLabel: '',
    modifiedSizes: [],
    modifyTotalSizeGB: 0,
    modifySizeRank: 0,
    modifyCountRank: 0,
    oldModified: createOldFiles(),

    folderCreateCount: 0,

    large: createLargeFiles(),

    items: [],
    summary: createBucketSummary( `Summary for ${userTitle}`, 'User' ),

    typesInfo: {
      count: 0,
      typeList: [],
      types: [],
      countRank: [],
      sizeRank: [],
    },
    
    duplicateInfo: {
      allNames: [],
      duplicateNames: [],
      duplicates: [],
      countRank: [],
      sizeRank: [],
      summary: createBucketSummary('Duplicate file info', 'Duplicate Files'),
    },

    folderInfo: {
      count: 0,
      folderRefs:[],
      folders: [],
      countRank: [],
      sizeRank: [],
    },

    uniqueInfo: {
      uniqueRolls: [],
      summary: createBucketSummary('Unique Permissions', 'Files with Unique Permissions'),
    },

    sharingInfo: {
      sharedItems: [],
      summary: createBucketSummary('Sharing Info', 'Shared Files'),
    },

  };

  return userSummary;

}

/***
 *    db    db d8888b. d8888b.  .d8b.  d888888b d88888b      d88888b d8888b. d888888b d888888b  .d88b.  d8888b. 
 *    88    88 88  `8D 88  `8D d8' `8b `~~88~~' 88'          88'     88  `8D   `88'   `~~88~~' .8P  Y8. 88  `8D 
 *    88    88 88oodD' 88   88 88ooo88    88    88ooooo      88ooooo 88   88    88       88    88    88 88oobY' 
 *    88    88 88~~~   88   88 88~~~88    88    88~~~~~      88~~~~~ 88   88    88       88    88    88 88`8b   
 *    88b  d88 88      88  .8D 88   88    88    88.          88.     88  .8D   .88.      88    `8b  d8' 88 `88. 
 *    ~Y8888P' 88      Y8888D' YP   YP    YP    Y88888P      Y88888P Y8888D' Y888888P    YP     `Y88P'  88   YD 
 *                                                                                                              
 *                                                                                                              
 */

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

/***
 *    db    db d8888b. d8888b.  .d8b.  d888888b d88888b       .d8b.  db    db d888888b db   db  .d88b.  d8888b. 
 *    88    88 88  `8D 88  `8D d8' `8b `~~88~~' 88'          d8' `8b 88    88 `~~88~~' 88   88 .8P  Y8. 88  `8D 
 *    88    88 88oodD' 88   88 88ooo88    88    88ooooo      88ooo88 88    88    88    88ooo88 88    88 88oobY' 
 *    88    88 88~~~   88   88 88~~~88    88    88~~~~~      88~~~88 88    88    88    88~~~88 88    88 88`8b   
 *    88b  d88 88      88  .8D 88   88    88    88.          88   88 88b  d88    88    88   88 `8b  d8' 88 `88. 
 *    ~Y8888P' 88      Y8888D' YP   YP    YP    Y88888P      YP   YP ~Y8888P'    YP    YP   YP  `Y88P'  88   YD 
 *                                                                                                              
 *                                                                                                              
 */

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

/***
 *     .o88b. d8888b. d88888b  .d8b.  d888888b d88888b      d8888b. db    db d8888b. db      d888888b  .o88b.  .d8b.  d888888b d88888b 
 *    d8P  Y8 88  `8D 88'     d8' `8b `~~88~~' 88'          88  `8D 88    88 88  `8D 88        `88'   d8P  Y8 d8' `8b `~~88~~' 88'     
 *    8P      88oobY' 88ooooo 88ooo88    88    88ooooo      88   88 88    88 88oodD' 88         88    8P      88ooo88    88    88ooooo 
 *    8b      88`8b   88~~~~~ 88~~~88    88    88~~~~~      88   88 88    88 88~~~   88         88    8b      88~~~88    88    88~~~~~ 
 *    Y8b  d8 88 `88. 88.     88   88    88    88.          88  .8D 88b  d88 88      88booo.   .88.   Y8b  d8 88   88    88    88.     
 *     `Y88P' 88   YD Y88888P YP   YP    YP    Y88888P      Y8888D' ~Y8888P' 88      Y88888P Y888888P  `Y88P' YP   YP    YP    Y88888P 
 *                                                                                                                                     
 *                                                                                                                                     
 */

export function createThisDuplicate ( detail : IItemDetail ) :IDuplicateFile {

  let iconInfo = getFileTypeIconInfo( detail.docIcon );

  let thisDup: IDuplicateFile = {
      name: detail.FileLeafRef,
      type: detail.docIcon, 
      locations: [],
      iconName: iconInfo.iconName,
      iconColor: iconInfo.iconColor,
      iconTitle: iconInfo.iconTitle,
      items: [],
      sizes: [],
      createdMs: [],
      modifiedMs: [],
      summary: createBucketSummary(`Dup: ${detail.FileLeafRef}`, 'Duplicate Files'),
      FileLeafRef: detail.FileLeafRef,
    };

  return thisDup;

}

/***
 *    db    db d8888b. d8888b.  .d8b.  d888888b d88888b      d8888b. db    db d8888b. db      d888888b  .o88b.  .d8b.  d888888b d88888b 
 *    88    88 88  `8D 88  `8D d8' `8b `~~88~~' 88'          88  `8D 88    88 88  `8D 88        `88'   d8P  Y8 d8' `8b `~~88~~' 88'     
 *    88    88 88oodD' 88   88 88ooo88    88    88ooooo      88   88 88    88 88oodD' 88         88    8P      88ooo88    88    88ooooo 
 *    88    88 88~~~   88   88 88~~~88    88    88~~~~~      88   88 88    88 88~~~   88         88    8b      88~~~88    88    88~~~~~ 
 *    88b  d88 88      88  .8D 88   88    88    88.          88  .8D 88b  d88 88      88booo.   .88.   Y8b  d8 88   88    88    88.     
 *    ~Y8888P' 88      Y8888D' YP   YP    YP    Y88888P      Y8888D' ~Y8888P' 88      Y88888P Y888888P  `Y88P' YP   YP    YP    Y88888P 
 *                                                                                                                                      
 *                                                                                                                                      
 */
export function updateThisDup ( thisDup: IDuplicateFile, detail : IItemDetail, LibraryUrl: string ) : IDuplicateFile {

  // title: string;
  // count: number;
  // size: number;
  // sizeGB: number;
  // sizeLabel: string;
  // countP: number;
  // sizeP: number;
  // sizeToCountRatio: number;  //Ratio of sizeP over countP.  Like 75% of all storage is filled by 5% of files ( 75/5 = 15 : 1 )
  // userTitles: string[];
  // userIds: number[];


  // thisDup.summary.count ++;
  // thisDup.summary.size += detail.size;

  thisDup.summary = updateBucketSummary( thisDup.summary, detail );


  // thisDup.summary.sizeGB = detail.size / 1e9;
  // thisDup.summary.sizeLabel = getSizeLabel( detail.size );
  // thisDup.summary.countP = 0;
  // thisDup.summary.sizeP = 0;
  // thisDup.summary.sizeToCountRatio = 0;
  if ( thisDup.summary.userTitles.indexOf( detail.authorTitle ) < 0 ) { thisDup.summary.userTitles.push( detail.authorTitle ) ; }
  if ( thisDup.summary.userTitles.indexOf( detail.editorTitle ) < 0 ) { thisDup.summary.userTitles.push( detail.editorTitle ) ; }

  if ( thisDup.summary.userIds.indexOf( detail.authorId ) < 0 ) { thisDup.summary.userIds.push( detail.authorId ) ; }
  if ( thisDup.summary.userIds.indexOf( detail.editorId ) < 0 ) { thisDup.summary.userIds.push( detail.editorId ) ; }

  thisDup.items.push( detail );
  thisDup.sizes.push(detail.size);

  thisDup.createdMs.push( detail.createMs ) ;
  thisDup.modifiedMs.push( detail.modMs ) ;

  // regex based on:  https://stackoverflow.com/a/17809074 and https://stackoverflow.com/a/494046
  // let replaceName = RegExp( /12(?![\s\S]*12)/ )
  let thisLocation = 'Unknown';
  if ( detail.FileLeafRef && detail.FileRef ){
    let lastIndex = detail.FileRef.lastIndexOf( detail.FileLeafRef );
    if ( lastIndex > 0 ) {
      thisLocation = detail.FileRef.substr(0, lastIndex );
      thisLocation = thisLocation.replace( LibraryUrl , ''); //Just show folder level url
    } else {
      debugger;
    }
  } else {
    debugger;
  }

  if ( thisDup.locations.indexOf(thisLocation ) < 0 ) { 
    thisDup.locations.push( thisLocation ) ; } 
  else { 
    debugger; 
  }


  return thisDup;

}

/***
 *     .o88b. d8888b. d88888b  .d8b.  d888888b d88888b      d888888b db    db d8888b. d88888b 
 *    d8P  Y8 88  `8D 88'     d8' `8b `~~88~~' 88'          `~~88~~' `8b  d8' 88  `8D 88'     
 *    8P      88oobY' 88ooooo 88ooo88    88    88ooooo         88     `8bd8'  88oodD' 88ooooo 
 *    8b      88`8b   88~~~~~ 88~~~88    88    88~~~~~         88       88    88~~~   88~~~~~ 
 *    Y8b  d8 88 `88. 88.     88   88    88    88.             88       88    88      88.     
 *     `Y88P' 88   YD Y88888P YP   YP    YP    Y88888P         YP       YP    88      Y88888P 
 *                                                                                            
 *                                                                                            
 */
export function createThisType ( docIcon: string ) :IFileType {

  let iconInfo = getFileTypeIconInfo( docIcon );

  let thisType: IFileType = {
    type: docIcon,
    iconName: iconInfo.iconName,
    iconColor: iconInfo.iconColor,
    iconTitle: iconInfo.iconTitle,
    avgSizeLabel: '',
    maxSizeLabel: '',
    avgSize: 0,
    maxSize: 0,
    items: [],
    sizes: [],
    createdMs: [],
    modifiedMs: [],
    summary: createBucketSummary('File Type info', 'File Type'),
  };

  return thisType;

}

/***
 *    db    db d8888b. d8888b.  .d8b.  d888888b d88888b      d888888b db    db d8888b. d88888b 
 *    88    88 88  `8D 88  `8D d8' `8b `~~88~~' 88'          `~~88~~' `8b  d8' 88  `8D 88'     
 *    88    88 88oodD' 88   88 88ooo88    88    88ooooo         88     `8bd8'  88oodD' 88ooooo 
 *    88    88 88~~~   88   88 88~~~88    88    88~~~~~         88       88    88~~~   88~~~~~ 
 *    88b  d88 88      88  .8D 88   88    88    88.             88       88    88      88.     
 *    ~Y8888P' 88      Y8888D' YP   YP    YP    Y88888P         YP       YP    88      Y88888P 
 *                                                                                             
 *                                                                                             
 */
export function updateThisType ( thisType: IFileType, detail : IItemDetail, ) : IFileType {

  thisType.items.push( detail );
  thisType.sizes.push(detail.size);

  thisType.createdMs.push( detail.createMs ) ;
  thisType.modifiedMs.push( detail.modMs ) ;
  thisType.summary = updateBucketSummary( thisType.summary, detail );

  return thisType;

}

/***
 *     .o88b. d8888b. d88888b  .d8b.  d888888b d88888b      d8888b.  .d8b.  d888888b  .o88b. db   db        d8888b.  .d8b.  d888888b  .d8b.  
 *    d8P  Y8 88  `8D 88'     d8' `8b `~~88~~' 88'          88  `8D d8' `8b `~~88~~' d8P  Y8 88   88        88  `8D d8' `8b `~~88~~' d8' `8b 
 *    8P      88oobY' 88ooooo 88ooo88    88    88ooooo      88oooY' 88ooo88    88    8P      88ooo88        88   88 88ooo88    88    88ooo88 
 *    8b      88`8b   88~~~~~ 88~~~88    88    88~~~~~      88~~~b. 88~~~88    88    8b      88~~~88 C8888D 88   88 88~~~88    88    88~~~88 
 *    Y8b  d8 88 `88. 88.     88   88    88    88.          88   8D 88   88    88    Y8b  d8 88   88        88  .8D 88   88    88    88   88 
 *     `Y88P' 88   YD Y88888P YP   YP    YP    Y88888P      Y8888P' YP   YP    YP     `Y88P' YP   YP        Y8888D' YP   YP    YP    YP   YP 
 *                                                                                                                                           
 *                                                                                                                                           
 */
//IBatchData, ILargeFiles, IUserFiles, IOldFiles
export function createBatchData ( currentUser: IUser, totalCount: number ):IBatchData {
  let currentUserId = currentUser ? currentUser.Id : 'TBD-Id';
  let currentUserTitle = currentUser ? currentUser.Title : 'TBD-Title';
  let currentUserName = !currentUser  ? 'TBD-Name' : currentUser.Name ? currentUser.Name  : currentUser.LoginName ;

  return {  
    totalCount: totalCount,
    summary: createBucketSummary('Duplicate file info', 'Batch'),
    significance: 0,
    isSignificant: false,
    items: [],
    typesInfo: {
      count: 0,
      typeList: [],
      types: [],
      countRank: [],
      sizeRank: [],
    },
    
    duplicateInfo: {
      allNames: [],
      duplicateNames: [],
      duplicates: [],
      countRank: [],
      sizeRank: [],
      summary: createBucketSummary('Duplicate file info', 'Duplicate Files'),
    },

    folderInfo: {
      count: 0,
      folderRefs:[],
      folders: [],
      countRank: [],
      sizeRank: [],
    },

    uniqueInfo: {
      uniqueRolls: [],
      summary: createBucketSummary('Unique Permissions', 'Files with Unique Permissions'),
    },

    large: createLargeFiles(),
    oldCreated: createOldFiles(),
    oldModified: createOldFiles(),
    
    userInfo: {

      count: 0,

      currentUser: createThisUser( null, currentUserId, currentUserTitle, currentUserName ),

      creatorIds: [],
      editorIds: [],
      allUsersIds: [],
      allUsers: [],

      createSizeRank: [],
      createCountRank: [],
      modifySizeRank: [],
      modifyCountRank: [],
    },

    sharingInfo: {
      sharedItems: [],
      summary: createBucketSummary('Sharing Info', 'Shared Files'),
    },

    analytics: {
      fetchMs: 0,
      analyzeMs: 0,
      fetchTime: null,
      fetchDuration: '',
      analyzeDuration: '',
      count: 0,
      msPerAnalyze: 0,
      msPerFetch: 0,
    }

  };
}

/***
 *    d8b   db  .d88b.  d888888b      db    db .d8888. d88888b d8888b. 
 *    888o  88 .8P  Y8. `~~88~~'      88    88 88'  YP 88'     88  `8D 
 *    88V8o 88 88    88    88         88    88 `8bo.   88ooooo 88   88 
 *    88 V8o88 88    88    88         88    88   `Y8b. 88~~~~~ 88   88 
 *    88  V888 `8b  d8'    88         88b  d88 db   8D 88.     88  .8D 
 *    VP   V8P  `Y88P'     YP         ~Y8888P' `8888Y' Y88888P Y8888D' 
 *                                                                     
 *                                                                     
 */
function createTypeRanks ( count: number ) : ITypeInfo {
  let theseInfos : ITypeInfo = {
    count: 0,
    countRank: [],
    sizeRank: [],
    typeList: [],
    types: [],
  };

  for (let index = 0; index < count; index++) {
    theseInfos.countRank.push( null );
    theseInfos.sizeRank.push( null );
  }

  return theseInfos;
}

function createDupRanks ( count: number ) : IDuplicateInfo {
  let theseInfos : IDuplicateInfo = {
    allNames: [],
    duplicates: [],
    duplicateNames: [],
    countRank: [],
    sizeRank: [],
    summary: createBucketSummary('Duplicate file info', 'Duplicate Files'),
  };

  for (let index = 0; index < count; index++) {
    theseInfos.countRank.push( null );
    theseInfos.sizeRank.push( null );
  }

  return theseInfos;
}

function createFolderRanks ( count: number ) : IFolderInfo {
  let theseInfos : IFolderInfo = {
    count: 0,
    folders: [],
    folderRefs:[],
    countRank: [],
    sizeRank: [],
  };

  for (let index = 0; index < count; index++) {
    theseInfos.countRank.push( null );
    theseInfos.sizeRank.push( null );
  }

  return theseInfos;
}

/***
 *     d888b  d88888b d888888b      d888888b d888888b d88888b .88b  d88. .d8888. 
 *    88' Y8b 88'     `~~88~~'        `88'   `~~88~~' 88'     88'YbdP`88 88'  YP 
 *    88      88ooooo    88            88       88    88ooooo 88  88  88 `8bo.   
 *    88  ooo 88~~~~~    88            88       88    88~~~~~ 88  88  88   `Y8b. 
 *    88. ~8~ 88.        88           .88.      88    88.     88  88  88 db   8D 
 *     Y888P  Y88888P    YP         Y888888P    YP    Y88888P YP  YP  YP `8888Y' 
 *                                                                               
 *                                                                               
 */
 export async function getStorageItems( pickedWeb: IPickedWebBasic , pickedList: IEXStorageList, fetchCount: number, currentUser: IUser, dataOptions: IDataOptions, addTheseItemsToState: any, setProgress: any, ) {

  // currentUser.Id = 466;  //REMOVE THIS LINE>>> USED FOR TESTING ONLY

  let webURL = pickedWeb.url;
  let listTitle = pickedList.Title;

  let items: any = null;
  let cleanedItems: IAllItemTypes[] = [];

  let isLoaded = false;

  let errMessage = '';
  let thisWebInstance = null;
  let createDateFromBatches: any[] = [];

  let batches: IEXStorageBatch[] = [];
 
  if ( fetchCount > 0 ) {
    try {
    
      // set the url for search
      // const searcher = Search(webURL);
  
      // This testing did not return anything I can understand that looks like a result.
      // this can accept any of the query types (text, ISearchQuery, or SearchQueryBuilder)
      // const results = await searcher(`Frauenhofer`);
      // console.log('Test searcher results', results);
  
      /***
       *                        .d8b.  db   d8b   db  .d8b.  d888888b d888888b 
       *           Vb          d8' `8b 88   I8I   88 d8' `8b   `88'   `~~88~~' 
       *            `Vb        88ooo88 88   I8I   88 88ooo88    88       88    
       *    C8888D    `V.      88~~~88 Y8   I8I   88 88~~~88    88       88    
       *              .d'      88   88 `8b d8'8b d8' 88   88   .88.      88    
       *            .dP        YP   YP  `8b8' `8d8'  YP   YP Y888888P    YP    
       *           dP                                                          
       *                                                                       
       */

      thisWebInstance = Web(webURL);
      let thisListObject = thisWebInstance.lists.getByTitle( listTitle );
      setProgress( 0 , pickedList.ItemCount, 'Getting ' + 'first' + ' batches of items' );
      try {
  
        let fetchStart = new Date();
        let startMs = fetchStart.getTime();
        let selectThese = listTitle === 'Preservation Hold Library' ? presHoldSelect : thisSelect;
        let expandThese = thisExpand;

        if ( dataOptions.getSharedDetails === true ) {
          selectThese =  [...selectThese, ...sharedWithSelect,  ];
          expandThese =  [...expandThese, ...sharedWithExpand,  ];
                  /**
           * This try just tries to get one item wtih the shared with details and if not, reverts back to baseline columns
           */
          try {
            items = await thisListObject.items.select(selectThese).expand(expandThese).top(1).filter('').getPaged();

          } catch (e){
            let helpfulErrorEnd = [ webURL, listTitle, null, null ].join('|');
            errMessage = getHelpfullErrorV2(e, false, true, [ 'BaseErrorTrace' , 'Failed', 'GetStorage ~ 59', helpfulErrorEnd ].join('|') );
            if ( errMessage.indexOf('SharedWithUsers') > -1 ) {
              //This library doesn't have SharedWithUsers.  Use Normal fetch
              selectThese = listTitle === 'Preservation Hold Library' ? presHoldSelect : thisSelect;
              expandThese = thisExpand;
            }
            errMessage = '';
          }
        }

        items = await thisListObject.items.select(selectThese).expand(expandThese).top(batchSize).filter('').getPaged(); 
  
        //Put basics into array just to check what order they are returned in.
        items.results.map( item => {
          let created = new Date(item.Created);
          let modified = new Date(item.Modified);
          let whichWasFirst = created.getTime() > modified.getTime() ? 'MOD' : 'Cre';
          let whichWasFirstDays = whichWasFirst + ' - '  + ( ( modified.getTime() - created.getTime() ) / msPerDay ).toPrecision(4);
          createDateFromBatches.push( { id: item.Id, FSOT: item.FileSystemObjectType, created: item.Created, modified: item.Modified, wwfd: whichWasFirstDays, wwf: whichWasFirst } );
        });

        batches = batches.concat( createThisBatch( items, startMs, 0 ) );
        for ( let i = 1; i < 150 ; i++ ) {
          let thisBatchStart = i * batchSize ;
          if ( items.hasNext && fetchCount > thisBatchStart ) {
            setProgress( thisBatchStart , fetchCount, `Fetching ${thisBatchStart} of ${ fetchCount } items` );
            fetchStart = new Date();
            startMs = fetchStart.getTime();
            items = await items.getNext();

            //Put basics into array just to check what order they are returned in.
            items.results.map( item => {
              let created = new Date(item.Created);
              let modified = new Date(item.Modified);
              let whichWasFirst = created.getTime() > modified.getTime() ? 'MOD' : 'Cre';
              let whichWasFirstDays = whichWasFirst + ' - ' + ( ( modified.getTime() - created.getTime() ) / msPerDay ).toPrecision(4);
              createDateFromBatches.push( { id: item.Id, FSOT: item.FileSystemObjectType, created: item.Created, modified: item.Modified, wwfd: whichWasFirstDays, wwf: whichWasFirst } );
            });


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

/***
 *                       .d8888. d88888b d888888b      db    db  .d8b.  d8888b. d888888b  .d8b.  d8888b. db      d88888b .d8888. 
 *           Vb          88'  YP 88'     `~~88~~'      88    88 d8' `8b 88  `8D   `88'   d8' `8b 88  `8D 88      88'     88'  YP 
 *            `Vb        `8bo.   88ooooo    88         Y8    8P 88ooo88 88oobY'    88    88ooo88 88oooY' 88      88ooooo `8bo.   
 *    C8888D    `V.        `Y8b. 88~~~~~    88         `8b  d8' 88~~~88 88`8b      88    88~~~88 88~~~b. 88      88~~~~~   `Y8b. 
 *              .d'      db   8D 88.        88          `8bd8'  88   88 88 `88.   .88.   88   88 88   8D 88booo. 88.     db   8D 
 *            .dP        `8888Y' Y88888P    YP            YP    YP   YP 88   YD Y888888P YP   YP Y8888P' Y88888P Y88888P `8888Y' 
 *           dP                                                                                                                  
 *                                                                                                                               
 */

  let batchData = createBatchData( currentUser, pickedList.ItemCount );
  //Add to large bucket
  let bigData = batchData.large;
  let oldData = batchData.oldCreated;

  let analyzeStart = new Date();
  let startMs2 = analyzeStart.getTime();

  //These variables are used to develop ceilings for metrics
  let largest: IItemDetail = null;
  let oldestCreate: IItemDetail = null;
  let oldestModified: IItemDetail = null;
  let userLargest: IItemDetail = null;
  let userOldestCreate: IItemDetail = null;
  let userOldestModified: IItemDetail = null;

  let allNames: string[] = [];
  let duplicateNames: string[] = [];
  let allNameItems: IDuplicateFile[] = [];

  let allFolderRefs: string [] = [];

  /***
 *                       .88b  d88.  .d8b.  d8888b.       .d8b.  db      db           d888888b d888888b d88888b .88b  d88. .d8888. 
 *           Vb          88'YbdP`88 d8' `8b 88  `8D      d8' `8b 88      88             `88'   `~~88~~' 88'     88'YbdP`88 88'  YP 
 *            `Vb        88  88  88 88ooo88 88oodD'      88ooo88 88      88              88       88    88ooooo 88  88  88 `8bo.   
 *    C8888D    `V.      88  88  88 88~~~88 88~~~        88~~~88 88      88              88       88    88~~~~~ 88  88  88   `Y8b. 
 *              .d'      88  88  88 88   88 88           88   88 88booo. 88booo.        .88.      88    88.     88  88  88 db   8D 
 *            .dP        YP  YP  YP YP   YP 88           YP   YP Y88888P Y88888P      Y888888P    YP    Y88888P YP  YP  YP `8888Y' 
 *           dP                                                                                                                    
 *                                                                                                                                 
 */

  batches.map( batch=> {

    let batchItems = dataOptions.getSharedDetails === true ? processSharedItems( batch.items ) : batch.items ;

    batchItems.map( ( item, itemIndex )=> {

      //Get item summary
      let detail: IItemDetail = createGenericItemDetail( batch.index , itemIndex, item, currentUser, dataOptions, pickedList.LibraryUrl );

      batchData.summary = updateBucketSummary( batchData.summary, detail );

      /***
       *                        d888b  d88888b d888888b       .d8b.  db    db d888888b db   db  .d88b.  d8888b.      d88888b d8888b. d888888b d888888b  .d88b.  d8888b. 
       *           Vb          88' Y8b 88'     `~~88~~'      d8' `8b 88    88 `~~88~~' 88   88 .8P  Y8. 88  `8D      88'     88  `8D   `88'   `~~88~~' .8P  Y8. 88  `8D 
       *            `Vb        88      88ooooo    88         88ooo88 88    88    88    88ooo88 88    88 88oobY'      88ooooo 88   88    88       88    88    88 88oobY' 
       *    C8888D    `V.      88  ooo 88~~~~~    88         88~~~88 88    88    88    88~~~88 88    88 88`8b        88~~~~~ 88   88    88       88    88    88 88`8b   
       *              .d'      88. ~8~ 88.        88         88   88 88b  d88    88    88   88 `8b  d8' 88 `88.      88.     88  .8D   .88.      88    `8b  d8' 88 `88. 
       *            .dP         Y888P  Y88888P    YP         YP   YP ~Y8888P'    YP    YP   YP  `Y88P'  88   YD      Y88888P Y8888D' Y888888P    YP     `Y88P'  88   YD 
       *           dP                                                                                                                                                   
       *                                                                                                                                                                
       */
      //Get index of authorId in array of all authorIds
      let createUserIndex = batchData.userInfo.creatorIds.indexOf( detail.authorId );
      if ( createUserIndex === -1 ) { 
        batchData.userInfo.creatorIds.push( detail.authorId  );
        createUserIndex = batchData.userInfo.creatorIds.length -1;
      }

      //Get index of editor in array of all editorIds
      let editUserIndex = batchData.userInfo.editorIds.indexOf( detail.editorId  );
      if ( editUserIndex === -1 ) { 
        batchData.userInfo.editorIds.push( detail.editorId  );
        editUserIndex = batchData.userInfo.editorIds.length -1;
      }

      //Get index of author in array of all allIds - to get the allUser Item for later use
      let createUserAllIndex = batchData.userInfo.allUsersIds.indexOf( detail.authorId );
      if ( createUserAllIndex === -1 ) { 
        batchData.userInfo.allUsersIds.push( detail.authorId  );
        batchData.userInfo.allUsers.push( createThisUser( detail, detail.authorId, detail.authorTitle, detail.authorShared )  );
        createUserAllIndex = batchData.userInfo.allUsers.length -1;
      }

      //Get index of editor in array of all allIds - to get the allUser Item for later use
      let editUserAllIndex = batchData.userInfo.allUsersIds.indexOf( detail.editorId  );
      if ( editUserAllIndex === -1 ) { 
        batchData.userInfo.allUsersIds.push( detail.editorId  );
        batchData.userInfo.allUsers.push( createThisUser( detail, detail.editorId, detail.editorTitle, detail.editorShared )  );
        editUserAllIndex = batchData.userInfo.allUsers.length -1;
      }

      batchData.userInfo.allUsers[ createUserAllIndex ] = updateThisAuthor( detail, batchData.userInfo.allUsers[ createUserAllIndex ]);
      batchData.userInfo.allUsers[ editUserAllIndex ] = updateThisEditor( detail, batchData.userInfo.allUsers[ editUserAllIndex ]);


      /***
       *                        d888b  d88888b d888888b      db       .d8b.  d8888b.  d888b  d88888b .d8888. d888888b       .d88b.  db      d8888b. d88888b .d8888. d888888b 
       *           Vb          88' Y8b 88'     `~~88~~'      88      d8' `8b 88  `8D 88' Y8b 88'     88'  YP `~~88~~'      .8P  Y8. 88      88  `8D 88'     88'  YP `~~88~~' 
       *            `Vb        88      88ooooo    88         88      88ooo88 88oobY' 88      88ooooo `8bo.      88         88    88 88      88   88 88ooooo `8bo.      88    
       *    C8888D    `V.      88  ooo 88~~~~~    88         88      88~~~88 88`8b   88  ooo 88~~~~~   `Y8b.    88         88    88 88      88   88 88~~~~~   `Y8b.    88    
       *              .d'      88. ~8~ 88.        88         88booo. 88   88 88 `88. 88. ~8~ 88.     db   8D    88         `8b  d8' 88booo. 88  .8D 88.     db   8D    88    
       *            .dP         Y888P  Y88888P    YP         Y88888P YP   YP 88   YD  Y888P  Y88888P `8888Y'    YP          `Y88P'  Y88888P Y8888D' Y88888P `8888Y'    YP    
       *           dP                                                                                                                                                        
       *                                                                                                                                                                     
       */
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

      /***
       *                       d8888b. db    db d888888b db      d8888b.      d888888b db    db d8888b. d88888b .d8888. 
       *           Vb          88  `8D 88    88   `88'   88      88  `8D      `~~88~~' `8b  d8' 88  `8D 88'     88'  YP 
       *            `Vb        88oooY' 88    88    88    88      88   88         88     `8bd8'  88oodD' 88ooooo `8bo.   
       *    C8888D    `V.      88~~~b. 88    88    88    88      88   88         88       88    88~~~   88~~~~~   `Y8b. 
       *              .d'      88   8D 88b  d88   .88.   88booo. 88  .8D         88       88    88      88.     db   8D 
       *            .dP        Y8888P' ~Y8888P' Y888888P Y88888P Y8888D'         YP       YP    88      Y88888P `8888Y' 
       *           dP                                                                                                   
       *                                                                                                                
       */
      //Build up Type list
      let typeIndex = batchData.typesInfo.typeList.indexOf( detail.docIcon );
      let typeIndexUser = batchData.userInfo.allUsers[ createUserAllIndex ].typesInfo.typeList.indexOf( detail.docIcon );

      if ( typeIndex < 0 ) {
        batchData.typesInfo.typeList.push( detail.docIcon );
        typeIndex = batchData.typesInfo.typeList.length - 1;
        batchData.typesInfo.types.push( createThisType(detail.docIcon) );
      }
      if ( typeIndexUser < 0 ) {
        batchData.userInfo.allUsers[ createUserAllIndex ].typesInfo.typeList.push( detail.docIcon );
        typeIndexUser = batchData.userInfo.allUsers[ createUserAllIndex ].typesInfo.typeList.length - 1;
        batchData.userInfo.allUsers[ createUserAllIndex ].typesInfo.types.push( createThisType(detail.docIcon) );
      }
      batchData.typesInfo.types[ typeIndex ] = updateThisType( batchData.typesInfo.types[ typeIndex ], detail );
      batchData.userInfo.allUsers[ createUserAllIndex ].typesInfo.types[ typeIndexUser ] = updateThisType( batchData.userInfo.allUsers[ createUserAllIndex ].typesInfo.types[ typeIndexUser ], detail );

      /***
       *    d8888b. db    db d888888b db      d8888b.      .d8888. db   db  .d8b.  d8888b. d888888b d8b   db  d888b       d888888b d8b   db d88888b  .d88b.  
       *    88  `8D 88    88   `88'   88      88  `8D      88'  YP 88   88 d8' `8b 88  `8D   `88'   888o  88 88' Y8b        `88'   888o  88 88'     .8P  Y8. 
       *    88oooY' 88    88    88    88      88   88      `8bo.   88ooo88 88ooo88 88oobY'    88    88V8o 88 88              88    88V8o 88 88ooo   88    88 
       *    88~~~b. 88    88    88    88      88   88        `Y8b. 88~~~88 88~~~88 88`8b      88    88 V8o88 88  ooo         88    88 V8o88 88~~~   88    88 
       *    88   8D 88b  d88   .88.   88booo. 88  .8D      db   8D 88   88 88   88 88 `88.   .88.   88  V888 88. ~8~        .88.   88  V888 88      `8b  d8' 
       *    Y8888P' ~Y8888P' Y888888P Y88888P Y8888D'      `8888Y' YP   YP YP   YP 88   YD Y888888P VP   V8P  Y888P       Y888888P VP   V8P YP       `Y88P'  
       *                                                                                                                                                     
       *                                                                                                                                                     
       */

      if ( detail.itemSharingInfo ) {
        batchData.sharingInfo.sharedItems.push( detail );
        batchData.sharingInfo.summary = updateBucketSummary( batchData.sharingInfo.summary, detail, );


      }

      /***
       *                       d8888b. db    db d888888b db      d8888b.      d8888b. db    db d8888b. db      d888888b  .o88b.  .d8b.  d888888b d88888b .d8888. 
       *           Vb          88  `8D 88    88   `88'   88      88  `8D      88  `8D 88    88 88  `8D 88        `88'   d8P  Y8 d8' `8b `~~88~~' 88'     88'  YP 
       *            `Vb        88oooY' 88    88    88    88      88   88      88   88 88    88 88oodD' 88         88    8P      88ooo88    88    88ooooo `8bo.   
       *    C8888D    `V.      88~~~b. 88    88    88    88      88   88      88   88 88    88 88~~~   88         88    8b      88~~~88    88    88~~~~~   `Y8b. 
       *              .d'      88   8D 88b  d88   .88.   88booo. 88  .8D      88  .8D 88b  d88 88      88booo.   .88.   Y8b  d8 88   88    88    88.     db   8D 
       *            .dP        Y8888P' ~Y8888P' Y888888P Y88888P Y8888D'      Y8888D' ~Y8888P' 88      Y88888P Y888888P  `Y88P' YP   YP    YP    Y88888P `8888Y' 
       *           dP                                                                                                                                            
       *                                                                                                                                                         
       */
      //Build up Duplicate list - only for filenames not folder names

      // let allNames: string[] = [];
      // let duplicateNames: string[] = [];

      if ( detail.isFolder !== true ) {
        let FileLeafRefLC = detail.FileLeafRef.toLowerCase();
        let dupIndex = allNames.indexOf( FileLeafRefLC );

        //Filename not encountered, add to All Names and create the Duplicate Item
        if ( dupIndex < 0 ) {
          allNames.push( FileLeafRefLC );
          dupIndex = allNames.length - 1;
          allNameItems.push( createThisDuplicate(detail) );
          allNameItems[ dupIndex ] = updateThisDup( allNameItems[ dupIndex ], detail, pickedList.LibraryUrl );

        //Filename was encountered, update the Duplicate Item
        } else {

          if ( duplicateNames.indexOf( FileLeafRefLC ) < 0 ) {
            duplicateNames.push( FileLeafRefLC ) ; 
          }

          allNameItems[ dupIndex ] = updateThisDup( allNameItems[ dupIndex ], detail, pickedList.LibraryUrl );
          batchData.duplicateInfo.summary = updateBucketSummary( batchData.duplicateInfo.summary, detail, );
        }

      }




      /***
       *                       d8888b. db    db d888888b db      d8888b.      db    db .d8888. d88888b d8888b. .d8888. 
       *           Vb          88  `8D 88    88   `88'   88      88  `8D      88    88 88'  YP 88'     88  `8D 88'  YP 
       *            `Vb        88oooY' 88    88    88    88      88   88      88    88 `8bo.   88ooooo 88oobY' `8bo.   
       *    C8888D    `V.      88~~~b. 88    88    88    88      88   88      88    88   `Y8b. 88~~~~~ 88`8b     `Y8b. 
       *              .d'      88   8D 88b  d88   .88.   88booo. 88  .8D      88b  d88 db   8D 88.     88 `88. db   8D 
       *            .dP        Y8888P' ~Y8888P' Y888888P Y88888P Y8888D'      ~Y8888P' `8888Y' Y88888P 88   YD `8888Y' 
       *           dP                                                                                                  
       *                                                                                                               
       */

      // if ( detail.currentUser === true ) { batchData.currentUser.items.push ( detail ) ; } 
      batchData.userInfo.allUsers[ createUserAllIndex ].items.push ( detail ) ;
      batchData.userInfo.allUsers[ createUserAllIndex ].summary = updateBucketSummary( batchData.userInfo.allUsers[ createUserAllIndex ].summary, detail );
      /***
       *                       db    db .d8888. d88888b d8888b.      d88888b  .d88b.  db      d8888b. d88888b d8888b. .d8888. 
       *           Vb          88    88 88'  YP 88'     88  `8D      88'     .8P  Y8. 88      88  `8D 88'     88  `8D 88'  YP 
       *            `Vb        88    88 `8bo.   88ooooo 88oobY'      88ooo   88    88 88      88   88 88ooooo 88oobY' `8bo.   
       *    C8888D    `V.      88    88   `Y8b. 88~~~~~ 88`8b        88~~~   88    88 88      88   88 88~~~~~ 88`8b     `Y8b. 
       *              .d'      88b  d88 db   8D 88.     88 `88.      88      `8b  d8' 88booo. 88  .8D 88.     88 `88. db   8D 
       *            .dP        ~Y8888P' `8888Y' Y88888P 88   YD      YP       `Y88P'  Y88888P Y8888D' Y88888P 88   YD `8888Y' 
       *           dP                                                                                                         
       *                                                                                                                      
       */
      batchData.folderInfo.folderRefs = allFolderRefs;
      let parentFolderIndex = allFolderRefs.indexOf( detail.parentFolder );
      let userParentFolderIndex = batchData.userInfo.allUsers[ createUserAllIndex ].folderInfo.folderRefs.indexOf( detail.parentFolder );

      if ( detail.isFolder === true ) { 
        //Create new IFolderDetail in all folders.
        allFolderRefs.push( detail.FileRef );
        parentFolderIndex = allFolderRefs.length -1;
        let folderAny : any = detail;
        let folderDetail : IFolderDetail = folderAny;
        folderDetail.sizeLabel = '';
        folderDetail.directItems = [];
        folderDetail.otherItems = [];
        folderDetail.totalCount = 0;
        folderDetail.totalSize = 0;
        folderDetail.directCount = 0;
        folderDetail.directSize = 0;
        folderDetail.directSizes = [];

        //push this new folder to top folder info
        batchData.folderInfo.folders.push ( folderDetail ) ;
        batchData.folderInfo.count ++;

        //Push new IFolderDetail in current user all folders.
        batchData.userInfo.allUsers[ createUserAllIndex ].folderInfo.folders.push ( folderDetail ) ;
        userParentFolderIndex = batchData.userInfo.allUsers[ createUserAllIndex ].folderInfo.folders.length - 1;

      } else { //This is not a folder but an item... update sizes
        if ( parentFolderIndex < 0 ) {
          console.log('WARNING - NOT ABLE TO FIND FOLDER:', detail.parentFolder );
        }
      }

      let thisDetailsParentFolder: IFolderDetail = batchData.folderInfo.folders[ parentFolderIndex ];
      if ( parentFolderIndex < 0 ) {
        console.log('WARNING - NOT ABLE TO FIND FOLDER:', detail.parentFolder );
      } else {

        // thisDetailsParentFolder.totalCount ++;
        // thisDetailsParentFolder.totalSize += detail.size;
      }


      /**
       * User Duplicates
       */

      /***
       *                       db    db .d8888. d88888b d8888b.      d8888b. d88888b d8888b. .88b  d88. .d8888. 
       *           Vb          88    88 88'  YP 88'     88  `8D      88  `8D 88'     88  `8D 88'YbdP`88 88'  YP 
       *            `Vb        88    88 `8bo.   88ooooo 88oobY'      88oodD' 88ooooo 88oobY' 88  88  88 `8bo.   
       *    C8888D    `V.      88    88   `Y8b. 88~~~~~ 88`8b        88~~~   88~~~~~ 88`8b   88  88  88   `Y8b. 
       *              .d'      88b  d88 db   8D 88.     88 `88.      88      88.     88 `88. 88  88  88 db   8D 
       *            .dP        ~Y8888P' `8888Y' Y88888P 88   YD      88      Y88888P 88   YD YP  YP  YP `8888Y' 
       *           dP                                                                                           
       *                                                                                                        
       */
      if ( detail.uniquePerms === true ) { 
        batchData.uniqueInfo.uniqueRolls.push ( detail ) ;
        batchData.userInfo.allUsers[ createUserAllIndex ].uniqueInfo.uniqueRolls.push ( detail ) ;
        batchData.uniqueInfo.summary = updateBucketSummary (batchData.uniqueInfo.summary , detail );
        batchData.userInfo.allUsers[ createUserAllIndex ].uniqueInfo.summary = updateBucketSummary (batchData.userInfo.allUsers[ createUserAllIndex ].uniqueInfo.summary , detail );
      }

      /***
       *                       db    db .d8888. d88888b d8888b.      .d8888. d888888b d88888D d88888b 
       *           Vb          88    88 88'  YP 88'     88  `8D      88'  YP   `88'   YP  d8' 88'     
       *            `Vb        88    88 `8bo.   88ooooo 88oobY'      `8bo.      88       d8'  88ooooo 
       *    C8888D    `V.      88    88   `Y8b. 88~~~~~ 88`8b          `Y8b.    88      d8'   88~~~~~ 
       *              .d'      88b  d88 db   8D 88.     88 `88.      db   8D   .88.    d8' db 88.     
       *            .dP        ~Y8888P' `8888Y' Y88888P 88   YD      `8888Y' Y888888P d88888P Y88888P 
       *           dP                                                                                 
       *                                                                                              
       */
      let userLarge = batchData.userInfo.allUsers[ createUserAllIndex ].large;
      if ( detail.size > 1e10 ) { 
        bigData.GT10G.push ( detail ) ;
        bigData.summary = updateBucketSummary (bigData.summary , detail );
        userLarge.GT10G.push ( detail ) ;
        userLarge.summary = updateBucketSummary (userLarge.summary , detail );

       } else if ( detail.size > 1e9 ) { 
        bigData.GT01G.push ( detail ) ; 
        bigData.summary = updateBucketSummary (bigData.summary , detail );
        batchData.userInfo.allUsers[ createUserAllIndex ].large.GT01G.push ( detail ) ;
        userLarge.summary = updateBucketSummary (userLarge.summary , detail );

      } else if ( detail.size > 1e8 ) { 
        bigData.GT100M.push ( detail ) ; 
        bigData.summary = updateBucketSummary (bigData.summary , detail );
        batchData.userInfo.allUsers[ createUserAllIndex ].large.GT100M.push ( detail ) ; 
        userLarge.summary = updateBucketSummary (userLarge.summary , detail );

      } else if ( detail.size > 1e7 ) { 
        bigData.GT10M.push ( detail ) ; 
        batchData.userInfo.allUsers[ createUserAllIndex ].large.GT10M.push ( detail ) ;

      }

      /***
       *                       db    db .d8888. d88888b d8888b.       .d8b.   d888b  d88888b 
       *           Vb          88    88 88'  YP 88'     88  `8D      d8' `8b 88' Y8b 88'     
       *            `Vb        88    88 `8bo.   88ooooo 88oobY'      88ooo88 88      88ooooo 
       *    C8888D    `V.      88    88   `Y8b. 88~~~~~ 88`8b        88~~~88 88  ooo 88~~~~~ 
       *              .d'      88b  d88 db   8D 88.     88 `88.      88   88 88. ~8~ 88.     
       *            .dP        ~Y8888P' `8888Y' Y88888P 88   YD      YP   YP  Y888P  Y88888P 
       *           dP                                                                        
       *                                                                                     
       */
      let theCurrentYear = getCurrentYear();

      let userOldCreated = batchData.userInfo.allUsers[ createUserAllIndex ].oldCreated;

      if ( detail.createYr < theCurrentYear - 4 ) { 
        oldData.Age5Yr.push ( detail ) ;
        oldData.summary = updateBucketSummary (oldData.summary , detail );
        batchData.userInfo.allUsers[ createUserAllIndex ].oldCreated.Age5Yr.push ( detail ) ;
        userOldCreated.summary = updateBucketSummary (userOldCreated.summary , detail );

       }

      else if ( detail.createYr < theCurrentYear - 3 ) { 
        oldData.Age4Yr.push ( detail ) ; 
        oldData.summary = updateBucketSummary (oldData.summary , detail );
        batchData.userInfo.allUsers[ createUserAllIndex ].oldCreated.Age4Yr.push ( detail ) ;
        userOldCreated.summary = updateBucketSummary (userOldCreated.summary , detail );

      }
      else if ( detail.createYr < theCurrentYear - 2 ) { 
        oldData.Age3Yr.push ( detail ) ; 
        oldData.summary = updateBucketSummary (oldData.summary , detail );
        batchData.userInfo.allUsers[ createUserAllIndex ].oldCreated.Age3Yr.push ( detail ) ;
        userOldCreated.summary = updateBucketSummary (userOldCreated.summary , detail );

      }
      else if ( detail.createYr < theCurrentYear - 1 ) { 
        oldData.Age2Yr.push ( detail ) ; 
        oldData.summary = updateBucketSummary (oldData.summary , detail );
        batchData.userInfo.allUsers[ createUserAllIndex ].oldCreated.Age2Yr.push ( detail ) ;
        userOldCreated.summary = updateBucketSummary (userOldCreated.summary , detail );

      }
      else if ( detail.createYr < theCurrentYear - 0 ) { 
        oldData.Age1Yr.push ( detail ) ; 
        userOldCreated.Age1Yr.push ( detail ) ;
      }

      let userOldModified = batchData.userInfo.allUsers[ editUserAllIndex ].oldModified;
      userOldModified.summary = updateBucketSummary( userOldModified.summary, detail );

      if ( detail.modYr < theCurrentYear - 4 ) { 
        batchData.oldModified.Age5Yr.push ( detail ) ;
        userOldModified.Age5Yr.push ( detail ) ;  
       }
      else if ( detail.modYr < theCurrentYear - 3 ) { 
        batchData.oldModified.Age4Yr.push ( detail ) ; 
        userOldModified.Age4Yr.push ( detail ) ;
      }
      else if ( detail.modYr < theCurrentYear - 2 ) { 
        batchData.oldModified.Age3Yr.push ( detail ) ; 
        userOldModified.Age3Yr.push ( detail ) ; 
      }
      else if ( detail.modYr < theCurrentYear - 1 ) { 
        batchData.oldModified.Age2Yr.push ( detail ) ; 
        userOldModified.Age2Yr.push ( detail ) ;
      }
      else if ( detail.modYr < theCurrentYear - 0 ) { 
        batchData.oldModified.Age1Yr.push ( detail ) ; 
        userOldModified.Age1Yr.push ( detail ) ;
      }


      
    /***
     *                       d88888b d8b   db d8888b.      d888888b d888888b d88888b .88b  d88.      .88b  d88.  .d8b.  d8888b. 
     *           Vb          88'     888o  88 88  `8D        `88'   `~~88~~' 88'     88'YbdP`88      88'YbdP`88 d8' `8b 88  `8D 
     *            `Vb        88ooooo 88V8o 88 88   88         88       88    88ooooo 88  88  88      88  88  88 88ooo88 88oodD' 
     *    C8888D    `V.      88~~~~~ 88 V8o88 88   88         88       88    88~~~~~ 88  88  88      88  88  88 88~~~88 88~~~   
     *              .d'      88.     88  V888 88  .8D        .88.      88    88.     88  88  88      88  88  88 88   88 88      
     *            .dP        Y88888P VP   V8P Y8888D'      Y888888P    YP    Y88888P YP  YP  YP      YP  YP  YP YP   YP 88      
     *           dP                                                                                                             
     *                                                                                                                          
     */
      cleanedItems.push( detail );

    });
  });


  cleanedItems.map ( detail => {
    let parentFolderIndex = allFolderRefs.indexOf( detail.parentFolder );

    let thisDetailsParentFolder: IFolderDetail = batchData.folderInfo.folders[ parentFolderIndex ];
    if ( parentFolderIndex < 0 ) {
      console.log('WARNING - NOT ABLE TO FIND FOLDER:', detail.parentFolder );
    } else {
  
      //Update main list of folder's stats for direct items
      if ( thisDetailsParentFolder.FileRef !== detail.FileRef ) {
        thisDetailsParentFolder.directSize += detail.size;
        // thisDetailsParentFolder.size += detail.size;
        thisDetailsParentFolder.directCount ++;
        thisDetailsParentFolder.directSizes.push( detail.size );
        thisDetailsParentFolder.directItems.push( detail );
      }

      // thisDetailsParentFolder.totalCount ++;
      // thisDetailsParentFolder.totalSize += detail.size;
    }
  });
  


  batchData.userInfo.count = batchData.userInfo.allUsersIds.length;

  /***
   *                       d88888b d888888b d8b   db d888888b .d8888. db   db      d888888b db    db d8888b. d88888b d888888b d8b   db d88888b  .d88b.  
   *           Vb          88'       `88'   888o  88   `88'   88'  YP 88   88      `~~88~~' `8b  d8' 88  `8D 88'       `88'   888o  88 88'     .8P  Y8. 
   *            `Vb        88ooo      88    88V8o 88    88    `8bo.   88ooo88         88     `8bd8'  88oodD' 88ooooo    88    88V8o 88 88ooo   88    88 
   *    C8888D    `V.      88~~~      88    88 V8o88    88      `Y8b. 88~~~88         88       88    88~~~   88~~~~~    88    88 V8o88 88~~~   88    88 
   *              .d'      88        .88.   88  V888   .88.   db   8D 88   88         88       88    88      88.       .88.   88  V888 88      `8b  d8' 
   *            .dP        YP      Y888888P VP   V8P Y888888P `8888Y' YP   YP         YP       YP    88      Y88888P Y888888P VP   V8P YP       `Y88P'  
   *           dP                                                                                                                                       
   *                                                                                                                                                    
   */
  //Update batchData typesInfo
  batchData.typesInfo.types.map( docType => {

    docType.avgSize = docType.summary.size/docType.summary.count;
    docType.maxSize = Math.max(...docType.sizes);
    docType.avgSizeLabel = docType.summary.count > 0 ? getSizeLabel(docType.avgSize) : '-';
    docType.maxSizeLabel = docType.summary.count > 0 ? getSizeLabel(docType.maxSize) : '-';

  });

  batchData.typesInfo.count = batchData.typesInfo.typeList.length;

  //Modify each user's typesInfo
  batchData.userInfo.allUsers.map( user => {
    user.typesInfo.types.map( docType => {
      docType.summary.sizeGB = docType.summary.size/1e9;
      docType.summary.sizeLabel = getSizeLabel( docType.summary.size );
      docType.summary.sizeP = docType.summary.size / user.createTotalSize * 100;
      docType.summary.countP = docType.summary.count / user.createCount * 100;
      docType.avgSize = docType.summary.size/docType.summary.count;
      docType.maxSize = Math.max(...docType.sizes);
      docType.avgSizeLabel = docType.summary.count > 0 ? getSizeLabel(docType.avgSize) : '-';
      docType.maxSizeLabel = docType.summary.count > 0 ? getSizeLabel(docType.maxSize) : '-';
    });
    user.typesInfo.count = user.typesInfo.typeList.length;
  });





  /***
   *                       d88888b d888888b d8b   db d888888b .d8888. db   db      db    db .d8888. d88888b d8888b. .d8888. 
   *           Vb          88'       `88'   888o  88   `88'   88'  YP 88   88      88    88 88'  YP 88'     88  `8D 88'  YP 
   *            `Vb        88ooo      88    88V8o 88    88    `8bo.   88ooo88      88    88 `8bo.   88ooooo 88oobY' `8bo.   
   *    C8888D    `V.      88~~~      88    88 V8o88    88      `Y8b. 88~~~88      88    88   `Y8b. 88~~~~~ 88`8b     `Y8b. 
   *              .d'      88        .88.   88  V888   .88.   db   8D 88   88      88b  d88 db   8D 88.     88 `88. db   8D 
   *            .dP        YP      Y888888P VP   V8P Y888888P `8888Y' YP   YP      ~Y8888P' `8888Y' Y88888P 88   YD `8888Y' 
   *           dP                                                                                                           
   *                                                                                                                        
   */
  //summarize Users data
  let allUserCreateSize: number[] = [];
  let allUserCreateCount: number[] = [];
  let allUserModifySize: number[] = [];
  let allUserModifyCount: number[] = [];

  batchData.userInfo.createSizeRank = expandArray( batchData.userInfo.count );
  batchData.userInfo.createCountRank = expandArray( batchData.userInfo.count );
  batchData.userInfo.modifySizeRank = expandArray( batchData.userInfo.count );
  batchData.userInfo.modifyCountRank = expandArray( batchData.userInfo.count );
  let userInfo = batchData.userInfo;
  
  // batchData.typesInfo = createTypeRanks( batchData.typesInfo.count );
  // let typeRanks = batchData.typeRanks;
  
  // batchData.duplicateRanks = createDupRanks( batchData.duplicateInfo.count );
  // let duplicateRanks = batchData.duplicateRanks;

  // batchData.folderRanks = createFolderRanks( batchData.folderInfo.count );
  // let folderRanks = batchData.folderRanks;

  batchData.userInfo.allUsers.map( user => {
    user.createTotalSizeGB = user.createTotalSize / 1e9;
    user.modifyTotalSizeGB = user.modifyTotalSize / 1e9;

    user.summary.size = user.createTotalSize;
    user.summary.count = user.createCount;
    user.summary.sizeGB = user.summary.size / 1e9;

    user.summary = updateBucketSummaryPercents( user.summary, batchData.summary );

    user.large.summary = updateBucketSummaryPercents( user.large.summary, user.summary );

    user.oldCreated.summary = updateBucketSummaryPercents( user.oldCreated.summary, user.summary );

    user.oldModified.summary = updateBucketSummaryPercents( user.oldModified.summary, user.summary );

    user.duplicateInfo.summary = updateBucketSummaryPercents( user.duplicateInfo.summary, user.summary );

    user.uniqueInfo.summary = updateBucketSummaryPercents( user.uniqueInfo.summary, user.summary );

    allUserCreateSize.push( user.createTotalSize );
    allUserCreateCount.push( user.createCount );
    allUserModifySize.push( user.modifyTotalSize );
    allUserModifyCount.push( user.modifyCount );

  });

  // batchData.userInfo.allUsers = sortObjectArrayByChildNumberKey( batchData.userInfo.allUsers, 'dec', 'summary.size');
  // batchData.userInfo.allUsers = sortObjectArrayByChildNumberKey( batchData.userInfo.allUsers, 'dec', 'summary.sizeToCountRatio');

  //Sort totals by largest first
  allUserCreateSize = sortNumberArray( allUserCreateSize , 'dec');
  allUserCreateCount = sortNumberArray( allUserCreateCount , 'dec');
  allUserModifySize = sortNumberArray( allUserModifySize , 'dec');
  allUserModifyCount = sortNumberArray( allUserModifyCount , 'dec');

  //Rank users based on all users counts and sizes
  batchData.userInfo.allUsers.map( ( user, userIndex) => {
    user.createSizeRank = allUserCreateSize.indexOf( user.createTotalSize );
    userInfo.createSizeRank = updateNextOpenIndex( userInfo.createSizeRank, user.createSizeRank, userIndex );

    user.createCountRank = allUserCreateCount.indexOf( user.createCount );
    userInfo.createCountRank = updateNextOpenIndex( userInfo.createCountRank, user.createCountRank, userIndex );

    user.modifySizeRank = allUserModifySize.indexOf( user.modifyTotalSize );
    userInfo.modifySizeRank = updateNextOpenIndex( userInfo.modifySizeRank, user.modifySizeRank, userIndex );

    user.modifyCountRank = allUserModifyCount.indexOf( user.modifyCount );
    userInfo.modifyCountRank = updateNextOpenIndex( userInfo.modifyCountRank, user.modifyCountRank, userIndex );

    user.createTotalSizeLabel = getSizeLabel( user.createTotalSize ); 
    user.modifyTotalSizeLabel = getSizeLabel( user.modifyTotalSize );

  });

  /***
   *                       d88888b d888888b d8b   db d888888b .d8888. db   db      d8888b. d888888b  d888b       d8888b.  .d8b.  d888888b  .d8b.  
   *           Vb          88'       `88'   888o  88   `88'   88'  YP 88   88      88  `8D   `88'   88' Y8b      88  `8D d8' `8b `~~88~~' d8' `8b 
   *            `Vb        88ooo      88    88V8o 88    88    `8bo.   88ooo88      88oooY'    88    88           88   88 88ooo88    88    88ooo88 
   *    C8888D    `V.      88~~~      88    88 V8o88    88      `Y8b. 88~~~88      88~~~b.    88    88  ooo      88   88 88~~~88    88    88~~~88 
   *              .d'      88        .88.   88  V888   .88.   db   8D 88   88      88   8D   .88.   88. ~8~      88  .8D 88   88    88    88   88 
   *            .dP        YP      Y888888P VP   V8P Y888888P `8888Y' YP   YP      Y8888P' Y888888P  Y888P       Y8888D' YP   YP    YP    YP   YP 
   *           dP                                                                                                                                 
   *                                                                                                                                              
   */

  bigData.summary = updateBucketSummaryPercents( bigData.summary, batchData.summary);
  oldData.summary = updateBucketSummaryPercents( oldData.summary, batchData.summary);
  batchData.duplicateInfo.summary = updateBucketSummaryPercents( batchData.duplicateInfo.summary, batchData.summary );
  batchData.uniqueInfo.summary = updateBucketSummaryPercents( batchData.uniqueInfo.summary, batchData.summary );

  batchData.folderInfo.folders.map( folder => {
    folder.sizeMB = folder.size / 1e6;
    folder.sizeLabel = getSizeLabel( folder.size );
  });

  /***
   *                       d88888b d888888b d8b   db d888888b .d8888. db   db      d8888b. db    db d8888b. db      d888888b  .o88b.  .d8b.  d888888b d88888b .d8888. 
   *           Vb          88'       `88'   888o  88   `88'   88'  YP 88   88      88  `8D 88    88 88  `8D 88        `88'   d8P  Y8 d8' `8b `~~88~~' 88'     88'  YP 
   *            `Vb        88ooo      88    88V8o 88    88    `8bo.   88ooo88      88   88 88    88 88oodD' 88         88    8P      88ooo88    88    88ooooo `8bo.   
   *    C8888D    `V.      88~~~      88    88 V8o88    88      `Y8b. 88~~~88      88   88 88    88 88~~~   88         88    8b      88~~~88    88    88~~~~~   `Y8b. 
   *              .d'      88        .88.   88  V888   .88.   db   8D 88   88      88  .8D 88b  d88 88      88booo.   .88.   Y8b  d8 88   88    88    88.     db   8D 
   *            .dP        YP      Y888888P VP   V8P Y888888P `8888Y' YP   YP      Y8888D' ~Y8888P' 88      Y88888P Y888888P  `Y88P' YP   YP    YP    Y88888P `8888Y' 
   *           dP                                                                                                                                                     
   *                                                                                                                                                                  
   */
  allNameItems.map( dup => {
    if ( dup.summary.count > 1 ) {
      batchData.duplicateInfo.duplicateNames.push( dup.name ) ;
      batchData.duplicateInfo.duplicates.push( dup ) ;
    }
  });

/***
 *                       d8888b. d88888b d888888b db    db d8888b. d8b   db      d888888b d8b   db d88888b  .d88b.  
 *           Vb          88  `8D 88'     `~~88~~' 88    88 88  `8D 888o  88        `88'   888o  88 88'     .8P  Y8. 
 *            `Vb        88oobY' 88ooooo    88    88    88 88oobY' 88V8o 88         88    88V8o 88 88ooo   88    88 
 *    C8888D    `V.      88`8b   88~~~~~    88    88    88 88`8b   88 V8o88         88    88 V8o88 88~~~   88    88 
 *              .d'      88 `88. 88.        88    88b  d88 88 `88. 88  V888        .88.   88  V888 88      `8b  d8' 
 *            .dP        88   YD Y88888P    YP    ~Y8888P' 88   YD VP   V8P      Y888888P VP   V8P YP       `Y88P'  
 *           dP                                                                                                     
 *                                                                                                                  
 */
  let analyzeEnd = new Date();
  let endMs2 = analyzeEnd.getTime();
  let analyzeMs = endMs2 - startMs2;

  let fetchMs = 0;
  let totalLength = 0;
  batches.map ( batch => { 
    fetchMs += batch.duration;
    totalLength += batch.items.length;
  });

  let currentUserAllIndex = batchData.userInfo.allUsersIds.indexOf( currentUser.Id );
  if ( currentUserAllIndex < 0 ) {
    //User was not created based on content... create a user profile in memory:
    let currentUserId = currentUser ? currentUser.Id : 'TBD-Id';
    let currentUserTitle = currentUser ? currentUser.Title : 'TBD-Title';
    let currentUserName = !currentUser  ? 'TBD-Name' : currentUser.Name ? currentUser.Name  : currentUser.LoginName ;

    let currentUserObj = createThisUser( null, currentUserId, currentUserTitle, currentUserName ) ;
    batchData.userInfo.count ++;

    currentUserObj.createSizeRank = batchData.userInfo.count - 1;
    currentUserObj.createCountRank = batchData.userInfo.count - 1;
    currentUserObj.modifyCountRank = batchData.userInfo.count - 1;
    currentUserObj.modifySizeRank = batchData.userInfo.count - 1;

    batchData.userInfo.allUsers.push( currentUserObj );
    batchData.userInfo.currentUser = currentUserObj;

    // batchData.userInfo.creatorIds.push( currentUserObj.userId );  //Not needed at this point
    // batchData.userInfo.editorIds.push( currentUserObj.userId );  //Not needed at this point
    batchData.userInfo.allUsersIds.push( currentUserObj.userId );

    batchData.userInfo.createSizeRank.push( batchData.userInfo.count - 1 );
    batchData.userInfo.createCountRank.push( batchData.userInfo.count - 1 );
    batchData.userInfo.modifySizeRank.push( batchData.userInfo.count - 1 );
    batchData.userInfo.modifyCountRank.push( batchData.userInfo.count - 1 );
    currentUserAllIndex = batchData.userInfo.allUsers.length - 1;

  }

  batchData.significance = batchData.summary.count > 0 ? batchData.summary.count / batchData.totalCount : 0 ;
  if ( batchData.significance > .95 ) { batchData.isSignificant = true ; }

  batchData.userInfo.currentUser = batchData.userInfo.allUsers [ currentUserAllIndex ];
  batchData.items = cleanedItems;

  let fetchTime = new Date();

  batchData.analytics = {
    fetchMs: fetchMs,
    analyzeMs: analyzeMs,
    fetchTime: fetchTime.toLocaleString(),
    fetchDuration: getCountLabel( fetchMs / ( 1000 * 60 ), 2 ) + ' minutes',
    analyzeDuration: ( analyzeMs / 1000 ).toFixed(4) + ' seconds',
    count: batchData.summary.count,
    msPerAnalyze: batchData.summary.count > 0 ?( analyzeMs / batchData.summary.count ) : null,
    msPerFetch: batchData.summary.count > 0 ?( fetchMs / batchData.summary.count ) : null,
  };

  let batchInfo = {
    batches: batches,
    batchData: batchData,
    totalLength: totalLength,
    userInfo: userInfo,
  };

  console.log('createDateFromBatches', createDateFromBatches );
  console.log('getStorageItems: fetchMs', fetchMs );
  console.log('getStorageItems: analyzeMs', analyzeMs );
  console.log('getStorageItems: totalLength', totalLength );
  console.log('getStorageItems: userInfo', userInfo );

  console.log('getStorageItems: batches', batches );
  console.log('getStorageItems: batchData', batchData );

  addTheseItemsToState( batchInfo );

  return { batches };
 
 }


 /***
 *     .o88b. d8888b. d88888b  .d8b.  d888888b d88888b      d888888b d888888b d88888b .88b  d88.        d8888b. d88888b d888888b  .d8b.  d888888b db      
 *    d8P  Y8 88  `8D 88'     d8' `8b `~~88~~' 88'            `88'   `~~88~~' 88'     88'YbdP`88        88  `8D 88'     `~~88~~' d8' `8b   `88'   88      
 *    8P      88oobY' 88ooooo 88ooo88    88    88ooooo         88       88    88ooooo 88  88  88        88   88 88ooooo    88    88ooo88    88    88      
 *    8b      88`8b   88~~~~~ 88~~~88    88    88~~~~~         88       88    88~~~~~ 88  88  88 C8888D 88   88 88~~~~~    88    88~~~88    88    88      
 *    Y8b  d8 88 `88. 88.     88   88    88    88.            .88.      88    88.     88  88  88        88  .8D 88.        88    88   88   .88.   88booo. 
 *     `Y88P' 88   YD Y88888P YP   YP    YP    Y88888P      Y888888P    YP    Y88888P YP  YP  YP        Y8888D' Y88888P    YP    YP   YP Y888888P Y88888P 
 *                                                                                                                                                        
 *                                                                                                                                                        
 */
 function createGenericItemDetail ( batchIndex:  number, itemIndex:  number, item: any, currentUser: IUser, dataOptions: IDataOptions, LibraryUrl: string ) : IItemDetail {
  let created = new Date(item.Created);
  let modified = new Date(item.Modified);

  let createYr = created.getFullYear();
  let modYr = modified.getFullYear();

  let isCurrentUser = item.AuthorId === currentUser.Id ? true : false;
  isCurrentUser = item.EditorId === currentUser.Id ? true : isCurrentUser;
  let size = item.FileSizeDisplay ? parseInt(item.FileSizeDisplay) : 0;

  let parentFolder =  item.FileRef.substring(0, item.FileRef.lastIndexOf('/') );
  let localFolder = `/${ parentFolder.replace( LibraryUrl, '' )}`;

  let authorSharedSplit = item.Author && item.Author.Name ? item.Author.Name.split('|') : [];
  let editorSharedSplit = item.Editor && item.Editor.Name ? item.Editor.Name.split('|') : [];

  let authorShared = authorSharedSplit.length > 0 ? authorSharedSplit[ 2 ].replace( domainEmail, '' ) : '';
  let editorShared = editorSharedSplit.length > 0 ? editorSharedSplit[ 2 ].replace( domainEmail, '' ) : '';

  let itemDetail: IItemDetail = {
    batch: batchIndex, //index of the batch in state.batches
    index: itemIndex, //index of item in state.batches[batch].items
    value: null, //value to highlight/sort for this detail
    Created: item.Created,
    Modified: item.Modified,
    created: created,
    modified: modified,
    authorId: item.AuthorId,
    editorId: item.EditorId,
    authorTitle: item.Author.Title,
    editorTitle: item.Editor.Title,
    authorName: item.Author.Name,
    authorShared: authorShared,
    editorShared: editorShared,
    editorName: item.Editor.Name,
    parentFolder: parentFolder,
    localFolder: localFolder,

    FileLeafRef: item.FileLeafRef,
    FileRef: item.FileRef,
    ServerRedirectedEmbedUrl: item.ServerRedirectedEmbedUrl,
    sizeLabel: getSizeLabel( size ),
    id: item.Id,
    currentUser: isCurrentUser,
    size: size,
    sizeMB: item.FileSizeDisplay ? Math.round( size / 1e6 * 100) / 100 : 0,
    createYr: createYr,
    modYr: modYr,
    bucket: `${createYr}-${modYr}`,
    createMs: created.getTime(),
    modMs: modified.getTime(),

    whichWasFirst: created.getTime() > modified.getTime() ? 'modfied' : 'created',
    whichWasFirstDays: ( ( modified.getTime() - created.getTime() ) / msPerDay ).toPrecision(4),

    ContentTypeId: item.ContentTypeId,
    ContentTypeName: '',
    docIcon: '',
    iconColor: '',
    iconName: '',
    iconTitle: '',
    version: item['OData__UIVersion'],
    versionlabel: item['OData__UIVersionString'],
    isMedia: false,
  };

  if ( item.CheckoutUserId ) { itemDetail.checkedOutId = item.CheckoutUserId; }
  if ( item.HasUniqueRoleAssignments ) { itemDetail.uniquePerms = item.HasUniqueRoleAssignments; }
  if ( item.FileSystemObjectType === 1 ) { itemDetail.isFolder = true; }

  if ( item.SharedWithDetails || item.SharedWithUsers || item.sharedEvents ) {
    itemDetail.itemSharingInfo = {
      sharedEvents: item.sharedEvents ? item.sharedEvents : [],
      SharedWithUsers: item.SharedWithUsers ? item.SharedWithUsers : [],
      FileRef: item.FileRef ,
      FileLeafRef: item.FileLeafRef ,
      FileSystemObjectType: item.FileSystemObjectType ,
      // SharedWithDetails: null,
    };
  }

  if ( dataOptions.useMediaTags === true ) {
    // itemDetail.MediaServiceAutoTags = item.MediaServiceAutoTags;
    // itemDetail.MediaServiceLocation = item.MediaServiceLocation;
    // itemDetail.MediaServiceOCR = item.MediaServiceOCR;
    // itemDetail.MediaServiceKeyPoints = item.MediaServiceKeyPoints;
    // itemDetail.MediaLengthInSeconds = item.MediaLengthInSeconds;
    ['MediaServiceAutoTags','MediaServiceLocation','MediaServiceOCR','MediaServiceKeyPoints','MediaLengthInSeconds'].map( key => {
      let keyProp = item[ key ];
      if ( keyProp && keyProp.length > 0 ) {  //Removed !== null because on WebPartDev Teams drag and drop it errored out.
        itemDetail[ key ] = keyProp;
        itemDetail.isMedia = true ;
      } else { itemDetail[ key ] = null ; }
    });
  }

  if ( item.DocIcon ) { 
    itemDetail.docIcon = item.DocIcon;

    let iconInfo = getFileTypeIconInfo( item.DocIcon );
    itemDetail.iconName = iconInfo.iconName; 
    itemDetail.iconColor = iconInfo.iconColor;   
    itemDetail.iconTitle = iconInfo.iconTitle;   

  } else if ( itemDetail.isFolder === true ) {

    itemDetail.docIcon = 'folder'; 
    itemDetail.iconName = 'OpenFolderHorizontal'; 
    itemDetail.iconColor = 'black';  
    itemDetail.iconTitle = 'Folder'; 
  }

  return itemDetail;

 }



 /***
 *     d888b  d88888b d888888b      d888888b  .o88b.  .d88b.  d8b   db      d888888b d8b   db d88888b  .d88b.  
 *    88' Y8b 88'     `~~88~~'        `88'   d8P  Y8 .8P  Y8. 888o  88        `88'   888o  88 88'     .8P  Y8. 
 *    88      88ooooo    88            88    8P      88    88 88V8o 88         88    88V8o 88 88ooo   88    88 
 *    88  ooo 88~~~~~    88            88    8b      88    88 88 V8o88         88    88 V8o88 88~~~   88    88 
 *    88. ~8~ 88.        88           .88.   Y8b  d8 `8b  d8' 88  V888        .88.   88  V888 88      `8b  d8' 
 *     Y888P  Y88888P    YP         Y888888P  `Y88P'  `Y88P'  VP   V8P      Y888888P VP   V8P YP       `Y88P'  
 *                                                                                                             
 *    import { getFileTypeIconInfo } from '@mikezimm/npmfunctions/dist/HelpInfo/Icons/stdECStorage';
 */
 function getFileTypeIconInfo( ext: string) {

  let iconColor = 'black';
  let iconName = ext;
  let iconTitle =  ext;
  switch (ext) {
    case 'xls':
    case 'xlsm':
    case 'xlsb':
    case 'xlsx':
      iconColor = 'darkgreen';
      iconName = 'ExcelDocument';
      break;

    case 'doc':
    case 'docx':
      iconColor = 'darkblue';
      iconName = 'WordDocument';
      break;

    case 'ppt':
    case 'pptx':
    case 'pptm':
      iconColor = 'firebrick';
      iconName = 'PowerPointDocument';
      break;

    case 'pdf':
      iconColor = 'red';
      break;

    case 'one':
    case 'onepkg':
      iconColor = 'purple';
      iconName = 'OneNoteLogo';
      break;

    case 'msg':
      iconColor = 'blue';
      iconName = 'OutlookLogo';
      break;

    case '7z':
    case 'zip':
      iconColor = 'blue';
      iconName = 'ZipFolder';
      break;

    case 'avi':
    case 'mp4':
    case 'wmf':
    case 'mov':
    case 'wmv':
      iconColor = 'dimgray';
      iconName = 'MSNVideosSolid';
      break;

    case 'msg':
      iconColor = 'blue';
      iconName = 'Microphone';
      break;

    case 'png':
    case 'gif':
    case 'jpg':
    case 'jpeg':
      iconColor = 'blue';
      iconName = 'Photo2';
      break;

    case 'txt':
    case 'csv':
      iconName = 'TextDocument';
      break;

    case 'dwg':
      iconName = 'PenWorkspace';
      break;

    default:
      iconName = 'FileTemplate';
      break;
  }

  return { iconName: iconName, iconColor: iconColor, iconTitle: iconTitle };

 }

 /***
 *     .o88b. d8888b. d88888b  .d8b.  d888888b d88888b      d888888b db   db d888888b .d8888.      d8888b.  .d8b.  d888888b  .o88b. db   db 
 *    d8P  Y8 88  `8D 88'     d8' `8b `~~88~~' 88'          `~~88~~' 88   88   `88'   88'  YP      88  `8D d8' `8b `~~88~~' d8P  Y8 88   88 
 *    8P      88oobY' 88ooooo 88ooo88    88    88ooooo         88    88ooo88    88    `8bo.        88oooY' 88ooo88    88    8P      88ooo88 
 *    8b      88`8b   88~~~~~ 88~~~88    88    88~~~~~         88    88~~~88    88      `Y8b.      88~~~b. 88~~~88    88    8b      88~~~88 
 *    Y8b  d8 88 `88. 88.     88   88    88    88.             88    88   88   .88.   db   8D      88   8D 88   88    88    Y8b  d8 88   88 
 *     `Y88P' 88   YD Y88888P YP   YP    YP    Y88888P         YP    YP   YP Y888888P `8888Y'      Y8888P' YP   YP    YP     `Y88P' YP   YP 
 *                                                                                                                                          
 *                                                                                                                                          
 */
 function createThisBatch ( results: any, start: number, batchIndex: number ) {
        
    let fetchEnd = new Date();
    let endMs = fetchEnd.getTime();
    let duration = endMs - start;
    let items = results.results;
    let count = items && items.length > 0 ? items.length : 0;
    let firstCreated = items && items.length > -1 ? new Date( items[0].Created ) : null;
    let lastCreated = items && items.length > -1 ? new Date( items[items.length - 1 ].Created ) : null;

    let batch: IEXStorageBatch = {
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

