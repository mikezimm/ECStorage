
import { IItemSharingInfo, ISharingEvent, ISharedWithUser } from './ISharingInterface';

import { sortNumberArray, sortObjectArrayByChildNumberKey } from '@mikezimm/npmfunctions/dist/Services/Arrays/sorting';

export const sharedWithSelect = [`SharedWithUsers/Title`,`SharedWithUsers/Name`,`SharedWithUsers/Id`,`SharedWithDetails`];
export const sharedWithExpand = ['SharedWithUsers'];

 export function processSharedItems( items: any[] ) {

  let domainEmail = window.location.hostname.replace('.sharepoint','');

  items.map( item => {

    if ( item.SharedWithDetails ) {
      if ( item.SharedWithUsers ) {
        item.SharedWithUsers.map( user => {
          delete user['odata.type'];  //Not needed
          delete user['odata.id'];  //Not needed
        });
      }

      item.SharedDetails = JSON.parse(item.SharedWithDetails);
      item.sharedEvents = Object.keys(item.SharedDetails).map( shareKey => {
        //This splits the Name prop which looks like this:  "Name":"i:0#.f|membership|first.lastName@tenant.com"
        let keys = shareKey.split('|');
        let detail = item.SharedDetails[ shareKey ];
        let SharedTime = getDateFromDetails( detail.DateTime );
        return {
          key: shareKey,
          keys: keys,
          sharedWith: keys[2] ? keys[2].replace( domainEmail, '') : 'Unknown',
          sharedBy: detail.LoginName  ? detail.LoginName.replace( domainEmail, '') : 'Unknown',
          DateTime: detail.DateTime,
          LoginName: detail.LoginName,
          TimeMS: SharedTime.getTime(),
          SharedTime: SharedTime,

          // Removed these items brought in from Pivot Tiles
          // GUID: item.GUID ,
          // odataEditLink: item.odataEditLink ,

          // AuthorId: item.AuthorId ,
          // Created: item.Created ,

          FileRef: item.FileRef ,
          FileLeafRef: item.FileLeafRef ,
          FileSystemObjectType: item.FileSystemObjectType ,
        
          // Modified: item.Modified ,
          // EditorId: item.EditorId ,
        
          // CheckoutUserId: item.CheckoutUserId ,

        };
      });

      item.sharedEvents = sortObjectArrayByChildNumberKey( item.sharedEvents, 'asc', 'TimeMS' );

    }
  });

  return items;

}



//SEND THIS TO npmFunctions

export function getDateFromDetails( details : string ) {

  let re = /-?\d+/; 
  let m = re.exec(details); 
  let d = new Date(parseInt(m[0]));
 
  return d;
 
 }