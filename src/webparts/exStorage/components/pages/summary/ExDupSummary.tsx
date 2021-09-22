import * as React from 'react';
import styles from '../../ExStorage.module.scss';

import { IExStorageState, IEXStorageList, IEXStorageBatch, IBatchData, IUserSummary, IFileType, ILargeFiles, IOldFiles, IDuplicateInfo } from '../../IExStorageState';
import { escape } from '@microsoft/sp-lodash-subset';

import {

} from "office-ui-fabric-react";

import { sortObjectArrayByNumberKey, sortNumberArray } from '@mikezimm/npmfunctions/dist/Services/Arrays/sorting';

import { getStorageItems, batchSize, createBatchData, getSizeLabel } from '../../ExFunctions';

import { createRatioNote } from './summaryFunctions';


export function createDupSummary ( dups: IDuplicateInfo, batchData: IBatchData ) : React.ReactElement {

  let currentDate = new Date();
  let currentYear = currentDate.getFullYear();

  let fullLoad = dups.summary.count === batchData.totalCount ? ' all' : ' ONLY';

  let loadPercent = batchData.totalCount !== 0 ? (( dups.summary.count / batchData.totalCount ) * 100) : 0;
  let loadPercentLabel = loadPercent.toFixed(1);
  let partialFlag = loadPercent === 100 ? '' : '*';

  let mainHeading = `Showing results for${fullLoad} ${ dups.summary.count } of ${ batchData.totalCount }`;
  let secondHeading = `This represents${fullLoad} ${ loadPercentLabel } of the files in this library.`;
  let tableRows = [];

  tableRows.push( <tr><td>{ `${ dups.summary.count } of ${ batchData.totalCount }`} </td><td>{ `Showing results for this many files in the library` }</td></tr> );
  tableRows.push( <tr><td>{ `or ${ loadPercentLabel }%`} </td><td>{ `% of all the files available` }</td></tr> );
  if ( loadPercent !== 100 ) {
    tableRows.push( <tr><td>{ partialFlag } </td><td>{ `Loading only part of the files may provide mis-leading results.` }</td></tr> );
    tableRows.push( <tr><td>{ null } </td><td>{ `For a complete picture, slide the Fetch counter all the way to the right and press Begin button` }</td></tr> );
  }
  tableRows.push( <tr><td>{ `${ dups.summary.sizeLabel } ${ partialFlag }`} </td><td>{ `Total size of all files oldFilesr created before ${ currentYear - 1 }`}</td></tr> );

  let GT100M = dups.summary.count;
  let GT100SizeLabel = getSizeLabel(dups.summary.size);

  tableRows.push( <tr><td>{ `${ GT100M } or ${ GT100SizeLabel } ${ partialFlag }`} </td><td>{ `Files created berfore ${ currentYear - 1 } ` }</td></tr> );
  tableRows.push( <tr><td>{ `<< Breaking News !! >>`} </td><td>{ createRatioNote( dups.summary,  '' ) }</td></tr> );

  let summaryTable = <table className={ styles.summaryTable }>
    { tableRows }
  </table>;

  const totalsInfo = <div className={ styles.flexWrapStart }>
    <div>{ mainHeading }</div>
    <div>{ secondHeading }</div>

  </div>;
  return <div style={{paddingTop: '20px' }}>
    { summaryTable }
  </div>;

}
