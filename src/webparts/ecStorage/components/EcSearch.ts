

/***
 *    d888888b .88b  d88. d8888b.  .d88b.  d8888b. d888888b       .d88b.  d88888b d88888b d888888b  .o88b. d888888b  .d8b.  db      
 *      `88'   88'YbdP`88 88  `8D .8P  Y8. 88  `8D `~~88~~'      .8P  Y8. 88'     88'       `88'   d8P  Y8   `88'   d8' `8b 88      
 *       88    88  88  88 88oodD' 88    88 88oobY'    88         88    88 88ooo   88ooo      88    8P         88    88ooo88 88      
 *       88    88  88  88 88~~~   88    88 88`8b      88         88    88 88~~~   88~~~      88    8b         88    88~~~88 88      
 *      .88.   88  88  88 88      `8b  d8' 88 `88.    88         `8b  d8' 88      88        .88.   Y8b  d8   .88.   88   88 88booo. 
 *    Y888888P YP  YP  YP 88       `Y88P'  88   YD    YP          `Y88P'  YP      YP      Y888888P  `Y88P' Y888888P YP   YP Y88888P 
 *                                                                                                                                  
 *                                                                                                                                  
 */
import { sp } from "@pnp/sp";

//https://sharepoint.stackexchange.com/questions/261222/spfx-and-pnp-sp-how-to-get-all-sites
//Just had to change SearchQuery to ISearchQuery.

import { ISearchQuery, SearchResults, ISearchResult } from "@pnp/sp/search";

import { IHubSiteWebData, IHubSiteInfo } from  "@pnp/sp/hubsites";
import "@pnp/sp/webs";
import "@pnp/sp/hubsites/web";

import { Web, IList, IItem } from "@pnp/sp/presets/all";

import { getHelpfullErrorV2 } from '@mikezimm/npmfunctions/dist/Services/Logging/ErrorHandler';

import { getPrincipalTypeString } from '@mikezimm/npmfunctions/dist/Services/Users/userServices';
import { getFullUrlFromSlashSitesUrl } from '@mikezimm/npmfunctions/dist/Services/Strings/urlServices';
import { IEcStorageState, IECStorageList, IECStorageBatch } from './IEcStorageState';
import { escape } from '@microsoft/sp-lodash-subset';

import { IPickedWebBasic, IPickedList, }  from '@mikezimm/npmfunctions/dist/Lists/IListInterfaces';

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

import { getExpandColumns, getKeysLike, getSelectColumns } from '@mikezimm/npmfunctions/dist/Lists/getFunctions';

import { getHelpfullError } from '@mikezimm/npmfunctions/dist/Services/Logging/ErrorHandler';


 
export interface MySearchResults extends ISearchResult {

}

export function getSearchedFiles( tenant: string, pickedList: IECStorageList,  ascSort :boolean  ) {

    //var departmentId = departmentId;
    // do a null check of department id
    //366df2ee-6476-4b15-a4fd-018dfae71e48 <= SPHub

    let pathEnd = pickedList ? pickedList.DocumentTemplateUrl.toLowerCase().indexOf('/forms/') : null;
    let path = tenant + pickedList ? pickedList.DocumentTemplateUrl.substr(0, pathEnd) : null;

    console.log('tenant:', tenant );
    console.log('path:', path );

    /**
     * FileSystemObjectType:  https://docs.microsoft.com/en-us/previous-versions/office/sharepoint-server/ee537053(v=office.15)#members
     *  File=0; Folder=1; Web=0
     * 
     * Source docs for testing:  https://docs.microsoft.com/en-us/sharepoint/technical-reference/query-variables
     * This works to get files in path and date range:  let query=`Path:"${path}"* AND Created>=2021-08-02 AND Created<2021-08-10` ;
     */

    let query=`Path:"${path}"* AND Created>=2021-08-02 AND Created<2021-08-10` ;  // ServerRelativePath:${path}* AND ContentClass:STS_ListItem AND Created:2020-09-07
    let thisSelect = ["*","Title", "ServerRelativeUrl", "ServerRelativePath", "ID", "Id", "Path", "Filename","FileLeafRef", "Author","Editor", 'Modified','Created','CheckoutUserId','HasUniqueRoleAssignments','FileSystemObjectType','FileSizeDisplay','FileLeafRef','LinkFilename','DocumentSummarySize','Size','SMTotalSize','File_x0020_Size','SMTotalFileStreamSize','FileSystemObjectType','OData__UIVersion','tp_UIVersion','OData__UIVersionString'];

    //Sort ascending by default
    let sortDirection = ascSort === false ? 1 : 0;

    /**
     *  Updated search query per pnpjs issue response:
     *  https://github.com/pnp/pnpjs/issues/1552#issuecomment-767837463
     * 
     * GET Managed properties here:  https://tenanat-admin.sharepoint.com/_layouts/15/searchadmin/ta_listmanagedproperties.aspx?level=tenant
     */
    sp.search(<ISearchQuery>{
          Querytext: query,
          SelectProperties: thisSelect,
          "RowLimit": 5000,
//          "StartRow": 0,
          "ClientType": "ContentSearchRegular",
          TrimDuplicates: false, //This is needed in order to also get the hub itself.
        })
          .then( ( res: SearchResults) => {
    
            console.log('Items from this list: ', res);
            console.log(res.RowCount);
            console.log(res.PrimarySearchResults);
            // entireResponse.hubs = res.PrimarySearchResults;

            // entireResponse.hubs.map( h => {
            //     h.sourceType = hubsCategory;
            // });
            // callback( entireResponse, custCategories, newData );

    });

    return;

}