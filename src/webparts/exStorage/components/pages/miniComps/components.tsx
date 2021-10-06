import * as React from 'react';
import styles from '../../ExStorage.module.scss';

import { IExStorageState, IEXStorageList, IEXStorageBatch, IBatchData, IUserSummary, IFileType, ILargeFiles, IOldFiles, IItemType, IIconArray, IItemDetail, IDuplicateFile } from '../../IExStorageState';

import { escape } from '@microsoft/sp-lodash-subset';

import {

} from "office-ui-fabric-react";

import { sortObjectArrayByChildNumberKey, sortNumberArray } from '@mikezimm/npmfunctions/dist/Services/Arrays/sorting';

import { getSizeLabel, getCommaSepLabel } from '@mikezimm/npmfunctions/dist/Services/Math/basicOperations';

import { Icon  } from 'office-ui-fabric-react/lib/Icon';

const flexWrapStart: React.CSSProperties = {
  display: 'flex',
  flexDirection: 'row',
  flexWrap: 'wrap',
  alignContent: 'center',
  justifyContent: 'flex-start',
  alignItems: 'center',
};

export function createItemsHeadingWithTypeIcons ( items: IItemDetail[] | IDuplicateFile[], itemsOrDups: IItemType, heading: any, icons: IIconArray[] ) {
  if ( !icons || icons.length === 0 ) {
    icons = createIconsArray( items );
  }
  let iconArray = createIconElementArray( icons );
  let element = 
    // <div className={styles.flexWrapStart}> //For some reason this did not work even though it was buried under the correct classname
    <div style={ flexWrapStart }>
      <h3>{ getCommaSepLabel( items.length ) } { itemsOrDups } found { heading }</h3> < div> { iconArray } </div>
    </div>;

  return element;
}

export function createIconsArray( itemsIn: IItemDetail[] | IDuplicateFile[] ) { // 
  let items: any[] = itemsIn;

  let icons: IIconArray[] = [];
  let iconNames: string[] = [];

  items.map( item => {
    let thisIcon = item.iconName ? item.iconName : 'Unknown';
    let idx = iconNames.indexOf( thisIcon);
    if ( idx < 0 ) {
      iconNames.push( thisIcon );
      icons.push( { iconColor: item.iconColor, iconName: item.iconName, iconTitle: item.iconTitle, sort1: item.size, sort2: 1 });
    } else { icons[ idx ].sort1 += item.size ;  icons[ idx ].sort2 ++ ; icons[ idx ].iconTitle = `${ item.iconTitle }: count: ${ getCommaSepLabel( icons[ idx ].sort2 ) }  size: ${ getSizeLabel( icons[ idx ].sort1 ) }` ; }
  });

  icons = sortObjectArrayByChildNumberKey( icons, 'dec', 'sort');

  return icons;

}

export function createIconElementArray( icons: IIconArray[] ) {

  let iconArray = icons.map( icon => {
    return ( <Icon iconName= { icon.iconName } title={ icon.iconTitle } style={ { fontSize: '24px', color: icon.iconColor, padding: '0px 0px 0px 15px', } }></Icon> );
  });

  return iconArray;

}