// import { IItemSharingInfo, ISharingEvent } from 'ISharingInterface';

import { IKnownMeta } from '../IExStorageState';

export interface ISharedWithUser {
  Title: string;
  Name: string;
  Id: number;
}

/**
 * This is details of one sharing event:  one person sharing with another person
 * It should be nexted under the item's IItemSharingInfo which has other item relavant info for later use like FileLeafRef
 * Should be nested under an item as an array of events
 */
export interface ISharingEvent {
  key: string;
  keys: string[];
  sharedWith: string;
  sharedBy: string;
  DateTime: string;
  TimeMS: number;
  LoginName: string;
  SharedTime: Date;
 
  //This is likely already in the parent but just adding it now so it's available when needed to build elements
  FileRef: string;
  FileLeafRef: string;
  FileSystemObjectType: number;


  ServerRedirectedEmbedUrl?: string;  //Only added after it's needed due to length
  parentFolder: string;  //Only added after it's needed due to length 

  id: number;
  iconName: string;
  iconColor: string;
  iconTitle: string;

  iconSearch: IKnownMeta; //Tried removing this but it caused issues with the auto-create title icons in Items.tsx so I'm adding it back.

  meta: IKnownMeta[];

  //Copying these down from item just for easier use.
  // GUID: string;
  // odataEditLink: string;
  // FileSystemObjectType: number;
  // AuthorId: number;
  // Created: string;
 
  // Modified: string;
  // EditorId: number;
 
  // CheckoutUserId: number;
 }
 
 /**
  * IItemSharingInfo is intended to be a sub-object on an item that has sharing.
  * It is used to then contain all the sharing details in a usable object
  */

 export interface IItemSharingInfo {
  SharedWithDetails?: string;
  // SharedDetails?: any;
  sharedEvents: ISharingEvent[];

  SharedWithUsers: ISharedWithUser[];

  // SharedWithUsersId?: number[];
  // Title: string;
 
  //Removed from interface derived from PivotTiles
  // Id: number;
  // ID: number;
  
  // GUID: string;
  // odataEditLink: string;
 
  // HasUniqueRoleAssignments: boolean;
  
  //These 3 are required to be used at this object level with components copied from Pivot Tiles

  FileRef: string;
  FileLeafRef: string;
 
  FileSystemObjectType: number;

  iconName: string;
  iconColor: string;
  iconTitle: string;

  iconSearch: IKnownMeta; //Tried removing this but it caused issues with the auto-create title icons in Items.tsx so I'm adding it back.
  meta: IKnownMeta[];
  id: number;  //Needed for on-click events
  
  // ServerRedirectedEmbedUrl: string;
  // ContentTypeId: string;
  // AuthorId: number;
  // Created: string;
 
  // Modified: string;
  // EditorId: number;
 
  // CheckoutUserId: number;
 }
 
 export interface IMySharingInfoSet {
   items: any[];
   elements: any[];
   isLoaded: boolean;
   errMessage: string;
  }
 
 export interface IMySharingInfo {
   history: IMySharingInfoSet;
   details: IMySharingInfoSet;
 }