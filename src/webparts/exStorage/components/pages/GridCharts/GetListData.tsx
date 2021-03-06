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

import { Web, IList, IItem } from "@pnp/sp/presets/all";

import "@pnp/sp/webs";
import "@pnp/sp/clientside-pages/web";
import "@pnp/sp/site-users/web";

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

import { addItemToArrayIfItDoesNotExist } from '@mikezimm/npmfunctions/dist/Services/Arrays/manipulation';

import { makeTheTimeObject,  } from '@mikezimm/npmfunctions/dist/Services/Time/timeObject';

import { getHelpfullError } from '@mikezimm/npmfunctions/dist/Services/Logging/ErrorHandler';

import { ensureUserInfo } from '@mikezimm/npmfunctions/dist/Services/Users/userServices';

import { getExpandColumns, getSelectColumns, IZBasicList, IPerformanceSettings, createFetchList, } from '@mikezimm/npmfunctions/dist/Lists/getFunctions';


 /***
 *    d888888b .88b  d88. d8888b.  .d88b.  d8888b. d888888b      db   db d88888b db      d8888b. d88888b d8888b. .d8888. 
 *      `88'   88'YbdP`88 88  `8D .8P  Y8. 88  `8D `~~88~~'      88   88 88'     88      88  `8D 88'     88  `8D 88'  YP 
 *       88    88  88  88 88oodD' 88    88 88oobY'    88         88ooo88 88ooooo 88      88oodD' 88ooooo 88oobY' `8bo.   
 *       88    88  88  88 88~~~   88    88 88`8b      88         88~~~88 88~~~~~ 88      88~~~   88~~~~~ 88`8b     `Y8b. 
 *      .88.   88  88  88 88      `8b  d8' 88 `88.    88         88   88 88.     88booo. 88      88.     88 `88. db   8D 
 *    Y888888P YP  YP  YP 88       `Y88P'  88   YD    YP         YP   YP Y88888P Y88888P 88      Y88888P 88   YD `8888Y' 
 *                                                                                                                       
 *                                                                                                                       
 */

 /***
 *    d888888b .88b  d88. d8888b.  .d88b.  d8888b. d888888b       .o88b.  .d88b.  .88b  d88. d8888b.  .d88b.  d8b   db d88888b d8b   db d888888b 
 *      `88'   88'YbdP`88 88  `8D .8P  Y8. 88  `8D `~~88~~'      d8P  Y8 .8P  Y8. 88'YbdP`88 88  `8D .8P  Y8. 888o  88 88'     888o  88 `~~88~~' 
 *       88    88  88  88 88oodD' 88    88 88oobY'    88         8P      88    88 88  88  88 88oodD' 88    88 88V8o 88 88ooooo 88V8o 88    88    
 *       88    88  88  88 88~~~   88    88 88`8b      88         8b      88    88 88  88  88 88~~~   88    88 88 V8o88 88~~~~~ 88 V8o88    88    
 *      .88.   88  88  88 88      `8b  d8' 88 `88.    88         Y8b  d8 `8b  d8' 88  88  88 88      `8b  d8' 88  V888 88.     88  V888    88    
 *    Y888888P YP  YP  YP 88       `Y88P'  88   YD    YP          `Y88P'  `Y88P'  YP  YP  YP 88       `Y88P'  VP   V8P Y88888P VP   V8P    YP    
 *                                                                                                                                               
 *                                                                                                                                               
 */

import { IGridItemInfo } from './IGridchartsState';


/***
 *    d88888b db    db d8888b.  .d88b.  d8888b. d888888b      d888888b d8b   db d888888b d88888b d8888b. d88888b  .d8b.   .o88b. d88888b .d8888. 
 *    88'     `8b  d8' 88  `8D .8P  Y8. 88  `8D `~~88~~'        `88'   888o  88 `~~88~~' 88'     88  `8D 88'     d8' `8b d8P  Y8 88'     88'  YP 
 *    88ooooo  `8bd8'  88oodD' 88    88 88oobY'    88            88    88V8o 88    88    88ooooo 88oobY' 88ooo   88ooo88 8P      88ooooo `8bo.   
 *    88~~~~~  .dPYb.  88~~~   88    88 88`8b      88            88    88 V8o88    88    88~~~~~ 88`8b   88~~~   88~~~88 8b      88~~~~~   `Y8b. 
 *    88.     .8P  Y8. 88      `8b  d8' 88 `88.    88           .88.   88  V888    88    88.     88 `88. 88      88   88 Y8b  d8 88.     db   8D 
 *    Y88888P YP    YP 88       `Y88P'  88   YD    YP         Y888888P VP   V8P    YP    Y88888P 88   YD YP      YP   YP  `Y88P' Y88888P `8888Y' 
 *                                                                                                                                               
 *                                                                                                                                               
 */


export interface IGridList extends IZBasicList {
    dropDownColumns: string[];
    dropDownSort: string[];
  }



//   d888b  d88888b d888888b  .d8b.  db      db      d888888b d888888b d88888b .88b  d88. .d8888. 
//  88' Y8b 88'     `~~88~~' d8' `8b 88      88        `88'   `~~88~~' 88'     88'YbdP`88 88'  YP 
//  88      88ooooo    88    88ooo88 88      88         88       88    88ooooo 88  88  88 `8bo.   
//  88  ooo 88~~~~~    88    88~~~88 88      88         88       88    88~~~~~ 88  88  88   `Y8b. 
//  88. ~8~ 88.        88    88   88 88booo. 88booo.   .88.      88    88.     88  88  88 db   8D 
//   Y888P  Y88888P    YP    YP   YP Y88888P Y88888P Y888888P    YP    Y88888P YP  YP  YP `8888Y' 
//                                                                                                
//        


// This is what it was before I split off the other part
export function updateAllItems( newItems: any[], fetchList: IGridList, addTheseItemsToState: any, setProgress: any, markComplete: any ) {
    let errMessage = null;

    let allItems : IGridItemInfo[] = newItems;

    /**
     * Add meta and searchString to every item
     */
    allItems.map( i => {
        //Add all date field objects
        fetchList.expandDates.map( d => {
            i['time' + d] = makeTheTimeObject(i[d]);
        });

        //Add Meta tags
        i.meta = buildMetaFromItem( i, fetchList );
        
        //Add Search string
        i.searchString = buildSearchStringFromItem( i, fetchList );
    });

    //private addTheseItemsToState( fetchList: IGridList, allItems , errMessage : string ) {
    allItems = addTheseItemsToState( fetchList, allItems, errMessage );

}



// This is what it was before I split off the other part
export async function getAllItems( fetchList: IGridList, addTheseItemsToState: any, setProgress: any, markComplete: any ): Promise<void>{

    let allItems : IGridItemInfo[] = [];
    let errMessage = '';

    let sourceUserInfo: any = null;
    try {
        sourceUserInfo = await ensureUserInfo( fetchList.webURL, fetchList.contextUserInfo.email );
    } catch (e) {
        errMessage = getHelpfullError(e, true, true);
    }


    fetchList.sourceUserInfo = sourceUserInfo;
    //lists.getById(listGUID).webs.orderBy("Title", true).get().then(function(result) {
    //let allItems : IGridItemInfo[] = await sp.web.webs.get();



    let thisListWeb = Web(fetchList.webURL);
    let selColumns = fetchList.selectColumnsStr;
    let expandThese = fetchList.expandColumnsStr;
    let staticCols = fetchList.staticColumns.length > 0 ? fetchList.staticColumns.join(',') : '';
    let selectCols = fetchList.minDataDownload === true ?  staticCols :  '*,' + staticCols;

    let thisListObject = thisListWeb.lists.getByTitle(fetchList.title);

    /**
     * IN FUTURE, ALWAYS BE SURE TO PUT SELECT AND EXPAND AFTER .ITEMS !!!!!!
     */

    try {
        let fetchCount = fetchList.fetchCount > 0 ? fetchList.fetchCount : 200;
        if ( fetchList.restFilter.length > 1 ) {
            allItems = await thisListObject.items.select(selectCols).expand(expandThese).orderBy('ID',false).top(fetchCount).filter(fetchList.restFilter).get();
        } else {
            allItems = await thisListObject.items.select(selectCols).expand(expandThese).orderBy('ID',false).top(fetchCount).get();
        }
    } catch (e) {
        errMessage = getHelpfullError(e, true, true);

    }

    /**
     * Add meta and searchString to every item
     */
    allItems.map( i => {
        //Add all date field objects
        fetchList.expandDates.map( d => {
            i['time' + d] = makeTheTimeObject(i[d]);
        });

        //Add Meta tags
        i.meta = buildMetaFromItem( i, fetchList );
        
        //Add Search string
        i.searchString = buildSearchStringFromItem( i, fetchList );
    });

    //private addTheseItemsToState( fetchList: IGridList, allItems , errMessage : string ) {
    allItems = addTheseItemsToState( fetchList, allItems, errMessage );

}







  
//  d8888b. db    db d888888b db      d8888b.      .88b  d88. d88888b d888888b  .d8b.  
//  88  `8D 88    88   `88'   88      88  `8D      88'YbdP`88 88'     `~~88~~' d8' `8b 
//  88oooY' 88    88    88    88      88   88      88  88  88 88ooooo    88    88ooo88 
//  88~~~b. 88    88    88    88      88   88      88  88  88 88~~~~~    88    88~~~88 
//  88   8D 88b  d88   .88.   88booo. 88  .8D      88  88  88 88.        88    88   88 
//  Y8888P' ~Y8888P' Y888888P Y88888P Y8888D'      YP  YP  YP Y88888P    YP    YP   YP 
//                                                                                     
//     

function buildMetaFromItem( theItem: IGridItemInfo, fetchList: IGridList, ) {
    let meta: string[] = theItem.meta ? theItem.meta : [];
    meta.push( 'All' );

    fetchList.metaColumns.map( c=> {
        if ( c.indexOf('/') > -1 ) { 
            let cols = c.split('/');
            //console.log( 'theItem', theItem);
            if ( theItem[ cols[0] ]) {
                meta = addItemToArrayIfItDoesNotExist( meta, theItem[ cols[0] ][ cols[1] ] ) ;
            } else { meta = addItemToArrayIfItDoesNotExist( meta, `. missing ${ c }` ) ; }
        } else if ( c.indexOf('.') > -1 ) { 
            let cols = c.split('.');
            if ( theItem[ cols[0] ]) {
                meta = addItemToArrayIfItDoesNotExist( meta, theItem[ cols[0] ][ cols[1]]  ) ;
            } else { meta = addItemToArrayIfItDoesNotExist( meta, `. missing ${ c }` ) ; }
        } else {
            meta = addItemToArrayIfItDoesNotExist( meta, theItem[ c ] ) ;
        }
        
    });

    fetchList.dropDownColumns.map( ( col , colIndex ) => {

        let actualColName = col.replace('>', '' ).replace('+', '' ).replace('-', '' );
        let parentColName = colIndex > 0 && col.indexOf('>') > -1 ? fetchList.dropDownColumns[colIndex - 1] : null;
        parentColName = parentColName !== null ? parentColName.replace('>', '' ).replace('+', '' ).replace('-', '' ) : null;

        let thisItemsChoices = theItem[ actualColName ];
        if ( parentColName !== null ) { thisItemsChoices = theItem[ parentColName ] + ' > ' + theItem[ actualColName ] ; }

        //2021-10-01:  Added this to solve the issue where CreatedBy and ModifiedBy are dropdowns and a person is in both columns.
        meta = addItemToArrayIfItDoesNotExist( meta, thisItemsChoices ) ;
        if ( theItem[ actualColName ] ) {
            meta = addItemToArrayIfItDoesNotExist( meta, col + '|>|' + theItem[ actualColName ] ) ;
        } else {
            meta = addItemToArrayIfItDoesNotExist( meta, col + '|>|-is Empty-' ) ;
        }

    });

    return meta;
}

//  d8888b. db    db d888888b db      d8888b.      .d8888. d88888b  .d8b.  d8888b.  .o88b. db   db 
//  88  `8D 88    88   `88'   88      88  `8D      88'  YP 88'     d8' `8b 88  `8D d8P  Y8 88   88 
//  88oooY' 88    88    88    88      88   88      `8bo.   88ooooo 88ooo88 88oobY' 8P      88ooo88 
//  88~~~b. 88    88    88    88      88   88        `Y8b. 88~~~~~ 88~~~88 88`8b   8b      88~~~88 
//  88   8D 88b  d88   .88.   88booo. 88  .8D      db   8D 88.     88   88 88 `88. Y8b  d8 88   88 
//  Y8888P' ~Y8888P' Y888888P Y88888P Y8888D'      `8888Y' Y88888P YP   YP 88   YD  `Y88P' YP   YP 
//                                                                                                 
//         

function buildSearchStringFromItem ( theItem: IGridItemInfo, fetchList: IGridList, ) {

    let result = '';
    let delim = '|||';

    if ( theItem.Title ) { result += 'Title=' + theItem.Title + delim ; }
    if ( theItem.Id ) { result += 'Id=' + theItem.Id + delim ; }

    fetchList.searchColumns.map( c => {
        let thisCol = c.replace('/','');
        if ( c.indexOf('/') > -1 ) { 
            let cols = c.split('/');
            if ( theItem[ cols[0] ] && theItem[ cols[0] ][ cols[1] ] ) { result += thisCol + '=' + theItem[ cols[0] ][ cols[1] ] + delim ; }
        } else if ( c.indexOf('.') > -1 ) { 
            let cols = c.split('.');
            if ( theItem[ cols[0] ] && theItem[ cols[0] ][ cols[1] ] ) { result += thisCol + '=' + theItem[ cols[0] ][ cols[1] ] + delim ; }
        } else {
            if ( theItem[thisCol] ) { result += thisCol + '=' + theItem[thisCol] + delim ; }
        }  

    });

    /**
     * Had this odata search from prior code:
     * if ( theItem['odata.type'] ) { result += theItem['odata.type'] + delim ; }
     * 
     */
    fetchList.odataSearch.map( odata => {
        if ( theItem[ odata ] ) { result += theItem[ odata ] + delim ; }
    });

    if ( theItem.meta.length > 0 ) { result += 'Meta=' + theItem.meta.join(',') + delim ; }

    result = result.toLowerCase();

    return result;

}