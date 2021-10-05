
import * as React from 'react';
import styles from '../../ExStorage.module.scss';

import { IExStorageState, IEXStorageList, IEXStorageBatch, IBatchData, IUserSummary, IFileType, IBucketSummary } from '../../IExStorageState';

import { getSizeLabel, getCommaSepLabel } from '@mikezimm/npmfunctions/dist/Services/Math/basicOperations';

export function createRatioNote ( summary: IBucketSummary, userLabel: string ) {
  if ( !userLabel || userLabel.length === 0 ) { userLabel = 'all' ; }
  return  `only ${ summary.countP.toPrecision(2) }% of ${ userLabel } files ( ${summary.count} ) account for ${ summary.sizeP.toPrecision(2) }%  ( ${summary.sizeLabel} ) of ${ userLabel } space`;
}

export function createTypeRatioNote ( summary: IFileType, userLabel: string ) {  //sizeToCountRatio
  if ( !userLabel || userLabel.length === 0 ) { userLabel = 'all' ; }
  let text = `${ summary.type }:  only ${ summary.countP.toPrecision(2) }% of ${ userLabel } files ( ${summary.count} )  account for ${ summary.sizeP.toPrecision(2) }% ( ${summary.sizeLabel} ) of ${ userLabel } space`;
  let title = `The size ( ${ summary.sizeLabel }) to count ( ${ summary.count }) ratio is ${ summary.sizeToCountRatio.toPrecision(2) }`;
  return  <span title={ title }>{ text }</span>;
}

export function createSummaryRangeRows ( tableRows: any[], summary: IBucketSummary ) {

  tableRows.push( <tr><td>{ summary.ranges.createRange } </td><td>{ `Old files CREATED during this timeframe` }</td></tr> );
  tableRows.push( <tr><td>{ summary.ranges.modifyRange } </td><td>{ `Old files MODIFIED during this timeframe` }</td></tr> );
  tableRows.push( <tr><td>{ summary.ranges.rangeAll } </td><td>{ `Old files were active during this timeframe` }</td></tr> );

  return tableRows;

}

export function createSummaryOldRows( tableRows: any[], summary: IBucketSummary, partialFlag: string ) {

  let currentDate = new Date();
  let currentYear = currentDate.getFullYear();

  tableRows.push( <tr><td>{ `${ summary.sizeLabel } ${ partialFlag }`} </td><td>{ `Total size of all files oldFilesr created before ${ currentYear - 1 }`}</td></tr> );

  let GT100M = getCommaSepLabel(summary.count);
  let GT100SizeLabel = getSizeLabel(summary.size);

  tableRows.push( <tr><td>{ `${ GT100M } or ${ GT100SizeLabel } ${ partialFlag }`} </td><td>{ `Files created before ${ currentYear - 1 } ` }</td></tr> );

  return tableRows;

}

export function createSummaryTopStats( tableRows: any[], summary: IBucketSummary, batchData: IBatchData, partialFlag: string ) {

  let loadPercentLabel = ( batchData.significance * 100 ).toFixed(1);

  tableRows.push( <tr><td>{ `${ getCommaSepLabel( summary.count) } of ${ getCommaSepLabel(batchData.totalCount) }`} </td><td>{ `Showing results for this many files in the library` }</td></tr> );
  tableRows.push( <tr><td>{ `or ${ loadPercentLabel }%`} </td><td>{ `% of all the files available` }</td></tr> );
  if ( batchData.significance !== 1 ) {
    tableRows.push( <tr><td>{ partialFlag } </td><td>{ `Loading only part of the files may provide mis-leading results.` }</td></tr> );
    tableRows.push( <tr><td>{ null } </td><td>{ `For a complete picture, slide the Fetch counter all the way to the right and press Begin button` }</td></tr> );
  }

  return tableRows;

}