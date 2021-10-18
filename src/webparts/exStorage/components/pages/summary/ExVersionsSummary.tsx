import * as React from 'react';
import styles from '../../ExStorage.module.scss';

import { IExStorageState, IEXStorageList, IEXStorageBatch, IBatchData, IUserSummary, IFileType, ILargeFiles, IOldFiles, IVersionInfo } from '../../IExStorageState';
import { escape } from '@microsoft/sp-lodash-subset';

import {

} from "office-ui-fabric-react";

import { sortObjectArrayByChildNumberKey, sortNumberArray } from '@mikezimm/npmfunctions/dist/Services/Arrays/sorting';

import { getSizeLabel, getCommaSepLabel } from '@mikezimm/npmfunctions/dist/Services/Math/basicOperations';

import { createRatioNote, createSummaryRangeRows, createSummaryOldRows, createSummaryTopStats, createOldModifiedRows, buildSummaryTable, createInfoRows, createSummaryLargeRows, createSummaryTypeRows } from './summaryFunctions';

export function createVersionsSummary ( versionInfo: IVersionInfo, batchData: IBatchData ) : React.ReactElement {

  let partialFlag = batchData.isSignificant === true ? '' : '*';

  let tableRows = [];

  tableRows.push( <tr><td>{ `NOTE:`} </td><td>Counts only include files but NOT folders</td></tr> );

  tableRows.push( <tr><td>{ `${ getCommaSepLabel( versionInfo.GT500.length ) } of ${ getCommaSepLabel(batchData.totalCount) }`} </td><td>Files with 500+ versions</td></tr> );
  tableRows.push( <tr><td>{ `${ getCommaSepLabel( versionInfo.GT100.length ) } of ${ getCommaSepLabel(batchData.totalCount) }`} </td><td>Files with 100 to 499 versions</td></tr> );
  tableRows.push( <tr><td>{ `${ getCommaSepLabel( versionInfo.draft.length ) } of ${ getCommaSepLabel(batchData.totalCount) }`} </td><td>Files with only draft versions &lt; 1.0</td></tr> );
  tableRows.push( <tr><td>{ `${ getCommaSepLabel( versionInfo.one.length ) } of ${ getCommaSepLabel(batchData.totalCount) }`} </td><td>Files only ONE version</td></tr> );
  tableRows.push( <tr><td>{ `${ getCommaSepLabel( versionInfo.GT1.length ) } of ${ getCommaSepLabel(batchData.totalCount) }`} </td><td>Files with 1.1 to 99 versions</td></tr> );


  let tableStyle = { marginTop: '1.4em'};

  return buildSummaryTable( tableRows, styles.exStorage, null ) ;

}
