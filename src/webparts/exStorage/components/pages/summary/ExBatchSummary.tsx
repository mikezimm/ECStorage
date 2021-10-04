import * as React from 'react';
import styles from '../../ExStorage.module.scss';

import { IExStorageState, IEXStorageList, IEXStorageBatch, IBatchData, IUserSummary, IFileType, ILargeFiles, IOldFiles } from '../../IExStorageState';
import { escape } from '@microsoft/sp-lodash-subset';

import {

} from "office-ui-fabric-react";

import { sortObjectArrayByChildNumberKey, sortNumberArray } from '@mikezimm/npmfunctions/dist/Services/Arrays/sorting';

import { getSizeLabel, getCommaSepLabel } from '@mikezimm/npmfunctions/dist/Services/Math/basicOperations';

import { createRatioNote } from './summaryFunctions';


export function createBatchSummary ( summary: IBatchData ) : React.ReactElement {
  let fullLoad = summary.count === summary.totalCount ? ' all' : ' ONLY';

  let loadPercent = summary.totalCount !== 0 ? (( summary.count / summary.totalCount ) * 100) : 0;
  let loadPercentLabel = loadPercent.toFixed(1);
  let partialFlag = loadPercent === 100 ? '' : '*';

  let mainHeading = `Showing results for${fullLoad} ${ summary.count } of ${ summary.totalCount }`;
  let secondHeading = `This represents${fullLoad} ${ loadPercentLabel } of the files in this library.`;
  let tableRows = [];

  tableRows.push( <tr><td>{ `${ getCommaSepLabel(summary.count) } of ${ getCommaSepLabel(summary.totalCount) }`} </td><td>{ `Showing results for this many files in the library` }</td></tr> );
  tableRows.push( <tr><td>{ `or ${ loadPercentLabel }%`} </td><td>{ `% of all the files available` }</td></tr> );
  if ( loadPercent !== 100 ) {
    tableRows.push( <tr><td>{ partialFlag } </td><td>{ `Loading only part of the files may provide mis-leading results.` }</td></tr> );
    tableRows.push( <tr><td>{ null } </td><td>{ `For a complete picture, slide the Fetch counter all the way to the right and press Begin button` }</td></tr> );
  }
  tableRows.push( <tr><td>{ `${ summary.sizeLabel } ${ partialFlag }`} </td><td>{ `Total size of all files fetched` }</td></tr> );
  tableRows.push( <tr><td>{ `${ summary.typesInfo.count } ${ partialFlag }`} </td><td>{ `File types found` }</td></tr> );
  tableRows.push( <tr><td>{ `${ summary.duplicateInfo.count } ${ partialFlag }`} </td><td>{ `Files that have more than one copy in the library` }</td></tr> );
  tableRows.push( <tr><td>{ `${ summary.uniqueInfo.count } ${ partialFlag }`} </td><td>{ `Folders/files with Unique Permissions` }</td></tr> );
  tableRows.push( <tr><td>{ `${ summary.userInfo.count } ${ partialFlag }`} </td><td>{ `Users who created/modified files` }</td></tr> );

  let GT100M = getCommaSepLabel(summary.large.summary.count);
  let GT100SizeLabel = getSizeLabel(summary.large.summary.size);

  tableRows.push( <tr><td>{ `${ GT100M } or ${ GT100SizeLabel } ${ partialFlag }`} </td><td>{ `Files larger than 100MB ` }</td></tr> );

  tableRows.push( <tr><td>{ `<< Breaking News !! >>`} </td><td>{ `only ${ createRatioNote( summary.large.summary,  '' ) }` }</td></tr> );

  let Age3YrCount = summary.oldModified.Age3Yr.length;
  Age3YrCount += summary.oldModified.Age4Yr.length;
  Age3YrCount += summary.oldModified.Age5Yr.length;

  tableRows.push( <tr><td>{ `${ Age3YrCount } ${ partialFlag }`} </td><td>{ `Files last modified more than a couple years ago` }</td></tr> );

  let summaryTable = <table className={ styles.summaryTable }>
    { tableRows }
  </table>;

  const totalsInfo = <div className={ styles.flexWrapStart }>
    <div>{ mainHeading }</div>
    <div>{ secondHeading }</div>

  </div>;
  return <div style={{  }}>
    { summaryTable }
  </div>;

}

