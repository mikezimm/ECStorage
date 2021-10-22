import * as React from 'react';

import {
    Icon,
  } from "office-ui-fabric-react";

import { Link } from 'office-ui-fabric-react';

import { sortObjectArrayByNumberKey, } from '@mikezimm/npmfunctions/dist/Services/Arrays/sorting';

import { IItemSharingInfo, ISharingEvent } from './ISharingInterface';
import { PopupWindowPosition } from '@microsoft/sp-property-pane';

import styles from './Sharing.module.scss';

import * as fpsAppIcons from '@mikezimm/npmfunctions/dist/Icons/standardEasyContents';

/**
 * This builds up the Shared History tab-page in Pivot Tiles library permissions
 * Provides Table of all individual sharing events on library in chronological order
 * @param sharedItems - like the array of items from a fetch
 * @param width 
 */
export function buildChronoSortedSharingEvents( sharedItems : any[], width: number ) {

    let allSharingEvents : ISharingEvent[] = [];
    let sharedElements: any[] = [];

    sharedItems.map( item => {
        if ( item.itemSharingInfo && item.itemSharingInfo.sharedEvents ) { 
            item.itemSharingInfo.sharedEvents.map( event => {
                allSharingEvents.push ( event );
            });
        }
    });
    //This sorts all the individual details by share timestamp
    allSharingEvents = sortObjectArrayByNumberKey( allSharingEvents, 'dec', 'TimeMS' );

    //This builds the elements based on the sorting
    allSharingEvents.map( share => {

        let sharedByName = share.sharedBy.split('@')[0];
        let sharedByDomain = sharedByName[1].split('.')[0] + '...';
        if ( share.sharedWith.indexOf( sharedByDomain ) > 0 ) { share.sharedWith = share.sharedBy.split('@')[0]; }

        let shortFileName = share.FileLeafRef && share.FileLeafRef.length > 0 ? share.FileLeafRef.substr(0,15) : '';
        if ( shortFileName.length < share.FileLeafRef.length ) { shortFileName += '...' ; }

        sharedElements.push( 
            <tr>
                <td> { share.SharedTime.toLocaleString() } </td>
                <td> { share.FileSystemObjectType === 0 ? 'File' : 'Folder' } </td>
                {/* <td> { share.GUID.split('-')[0] + '...' } </td> */}
                <td title={ share.FileLeafRef }> { <Link onClick={ openLinkInNewTabUsingDatahref } data-href= { share.FileRef }>{ shortFileName }</Link> } </td>
                <td> { sharedByName } </td>
                <td> { share.sharedWith } </td>

            </tr>
          );
    });

    return sharedElements;

}

/**
 * This builds up the Shared Details tab-page in Pivot Tiles library permissions
 * This shows all sharing grouped by the file that was shared
 * @param sharedItems
 * @param width 
 */
export function buildShareEventsGroupedByItem( sharedItems : any[], width: number ) {

    let sharedElements: any[] = [];


    //This builds the elements based on the sorting
    sharedItems.map( ( item, index )  => {

        let shortFileName = item.FileLeafRef && item.FileLeafRef.length > 0 ? item.FileLeafRef.substr(0,25) : '';
        if ( shortFileName.length < item.FileLeafRef.length ) { shortFileName += '...' ; }

        const UniquePermIcon: JSX.Element = <div id={ index.toString() } > { fpsAppIcons.UniquePerms } </div>;

        let shareTable = createDetailsShareTable( item, false, true, '100%Wide' );

        sharedElements.push( 
            <tr>
                <td > { UniquePermIcon } </td>
                <td> { item.FileSystemObjectType === 0 ? 'File' : 'Folder' } </td>
                {/* <td> { share.GUID.split('-')[0] + '...' } </td> */}
                <td title={ item.FileLeafRef }> { <Link onClick={ openLinkInNewTabUsingDatahref } data-href= { item.FileRef }>{ shortFileName }</Link> } </td>
                <td> { shareTable } </td>

            </tr>
          );
    });

    return sharedElements;

}

 /**
 * This just creates the 3 column table for each file/item showing When, who shared, with whome.
 * Can be consumed as a cell in a larger table of all shared files or just for a specific file.
  * @param item 
  * @param headings 
  * @param cleanCells - this will remove Date and Shared By if both of those are the same as the previous row.
  * @param tableStyle 
  */
export function createDetailsShareTable( item: any, headings: boolean, cleanCells: boolean, tableStyle: 'pad30' | '100%Wide' ) {

    let hasSharing = item.itemSharingInfo && item.itemSharingInfo.sharedEvents ? true : false;

    if ( hasSharing === false ) { return null; }

    let firstShareDateMS = 3618105359201;
    let lastShareDateMS = 0;

    let firstShareDate = null;
    let lastShareDate = null;

    let sharedByPeopleArray = [];
    let thisFileShares = [];

    let itemSharingInfo: IItemSharingInfo = item.itemSharingInfo;

    if ( itemSharingInfo && itemSharingInfo.sharedEvents ) {
        itemSharingInfo.sharedEvents.map( ( event, index ) =>{
            let lastEvent = index > 0 ? itemSharingInfo.sharedEvents[ index - 1 ] : null;
            let isSameAsLast = index > 0 && event.sharedBy === lastEvent.sharedBy && event.TimeMS === lastEvent.TimeMS ? true : false;

            let sharedByName = event.sharedBy.split('@')[0];
            let sharedByDomain = sharedByName[1].split('.')[0] + '...';
            if ( event.sharedWith.indexOf( sharedByDomain ) > 0 ) { event.sharedWith = event.sharedBy.split('@')[0]; }
    
            if ( event.TimeMS > lastShareDateMS ) { lastShareDate = event.SharedTime; lastShareDateMS = event.TimeMS ; }
            if ( event.TimeMS < firstShareDateMS ) { firstShareDate = event.SharedTime; firstShareDateMS = event.TimeMS ; }
            sharedByPeopleArray.push( event.sharedWith );
    
            thisFileShares.push( 
                <tr>
                    <td> { isSameAsLast === true ? '...' : event.SharedTime.toLocaleString() } </td>
                    <td> { isSameAsLast === true ? '...' : sharedByName } </td>
                    <td> { event.sharedWith } </td>
                </tr>
              );
    
        });
    }
    
    let shareTimeFrame = firstShareDate !== null ? firstShareDate.toLocaleString() : null;
    if ( lastShareDate !== null && firstShareDateMS !== lastShareDateMS ) { shareTimeFrame += ' - ' + lastShareDate.toLocaleString() ;  }

    let cellClass = tableStyle === 'pad30' ? styles.padAllCellsLeft : null ;
    let tableWidth = tableStyle === '100%Wide' ? '100%' : null ;

    let headingRow = headings !== true ? null : <tr>
        <th>Date</th>
        <th>Shared By</th>
        <th>Shared With</th>
    </tr>;

    let shareTable = thisFileShares.length === 0 ? null : <table className={ cellClass } style={{ width: tableWidth }}>
        { headingRow }
        { thisFileShares }
    </table>;

    return shareTable;

}

// function handleClickOnLink(ev: React.MouseEvent<unknown>) {
function openLinkInNewTabUsingDatahref( e: any ) {
    e.preventDefault();
    let testElement = e.nativeEvent.target;
    const href = testElement.getAttribute('data-href');
    window.open( href, '_blank' );
  }

  export function buildConstructionElement( mainContent: any, additionalContent: any ) {
      
    let iconStyles: any = { root: {
        fontSize: 'larger',
        // fontWeight: 700,
        color: 'red',
        // paddingRight: '30px',
        // paddingLeft: '30px',
    }};

    const leftIcon = <Icon iconName={'ConstructionCone'} styles = {iconStyles}/>;
    const rightIcon = <Icon iconName={'ConstructionConeSolid'} styles = {iconStyles}/>;

    let element = <div style={{ padding: '5px 50px 30px 50px'}}>
        <div style={{ fontSize: 'x-large', paddingBottom: '5px', textAlign: 'center' }}>
        { leftIcon } <div style={{display: 'inline-block', padding: '0 30px'}}> { mainContent } </div> { rightIcon } 
        { additionalContent }
        </div>
    </div>;

    return element;

  }