import * as React from 'react';
import styles from '../../ExStorage.module.scss';
import stylesMini from './mini.module.scss';

import {
  TooltipHost, ITooltipHostStyles
} from "office-ui-fabric-react";

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

export function createItemsHeadingWithTypeIcons ( items: IItemDetail[] | IDuplicateFile[], itemType: IItemType, heading: any, tooltip: string, icons: IIconArray[], onClickIcon: any ) {
  if ( !icons || icons.length === 0 ) {
    icons = createIconsArray( items );
  }
  let iconArray = createIconElementArray( icons, onClickIcon );
  const calloutProps = { gapSpace: 0 };

  let headingSpan = heading;

  if ( tooltip && tooltip.length > 0 ) {
    headingSpan = <TooltipHost content={ tooltip } id={'tooltipsearchbox'} calloutProps={calloutProps}>
      { heading }
    </TooltipHost>;
  }

  let element = 
    // <div className={styles.flexWrapStart}> //For some reason this did not work even though it was buried under the correct classname
      <div style={ flexWrapStart }>
        <h3>{ getCommaSepLabel( items.length ) } { itemType } found { headingSpan }</h3> < div> { iconArray } </div>
      </div>;

  return element;
}

/**
 * createIconsArray creates an array of icon objects that can be passed to createIconElementArray to create elements.
 * However, it also gets some extra data like an on-hover title showing file type stats.
 * @param itemsIn 
 */
export function createIconsArray( itemsIn: IItemDetail[] | IDuplicateFile[] ) { // 
  let items: any[] = itemsIn;

  let icons: IIconArray[] = [];
  let iconNames: string[] = [];

  items.map( item => {
    let thisIcon = item.iconName ? item.iconName : 'Unknown';
    let idx = iconNames.indexOf( thisIcon);
    if ( idx < 0 ) {
      iconNames.push( thisIcon );
      let iconTitle = `${ item.iconTitle }: count: ${ 1 }  size: ${ getSizeLabel(item.size ) }`;
      icons.push( { iconColor: item.iconColor, iconName: item.iconName, iconTitle: iconTitle, iconSearch: item.iconSearch, sort1: item.size, sort2: 1 });
    } else { icons[ idx ].sort1 += item.size ;  icons[ idx ].sort2 ++ ; icons[ idx ].iconTitle = `${ item.iconTitle }: count: ${ getCommaSepLabel( icons[ idx ].sort2 ) }  size: ${ getSizeLabel( icons[ idx ].sort1 ) }` ; }
  });

  icons = sortObjectArrayByChildNumberKey( icons, 'dec', 'sort1');

  return icons;

}

export function createIconElementArray( icons: IIconArray[], onClickIcon: any ) {

  let iconArray = icons.map( icon => {
    return ( <Icon iconName= { icon.iconName } id={ icon.iconSearch } data-search={ icon.iconSearch } title={ icon.iconTitle } onClick= { onClickIcon } style={ { fontSize: '24px', color: icon.iconColor, padding: '0px 0px 0px 15px', } }></Icon> );
  });

  return iconArray;

}


export function nothingToShow( sub?: string, head?: string ) {

  let heading = head ? head : 'Well, not sure how to tell you this but I can\'t find anything in this category :(';
  let subHeading = sub ? sub : '';
  let element = <div className={ stylesMini.noItems }>
    <h2>{ heading }</h2>
    { subHeading === '' ? null  : <h4>{ subHeading }</h4> }
  </div>;

  return element;

} 