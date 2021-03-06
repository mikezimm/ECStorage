import * as React from 'react';
import styles from '../../ExStorage.module.scss';

import { IExStorageState, IEXStorageList, IEXStorageBatch, IBatchData, IUserSummary, IFileType, ILargeFiles, IOldFiles } from '../../IExStorageState';
import { escape } from '@microsoft/sp-lodash-subset';

import {

} from "office-ui-fabric-react";

import { sortObjectArrayByChildNumberKey, sortNumberArray } from '@mikezimm/npmfunctions/dist/Services/Arrays/sorting';

import { getSizeLabel, getCommaSepLabel } from '@mikezimm/npmfunctions/dist/Services/Math/basicOperations';

import { createRatioNote, createSummaryRangeRows, createSummaryOldRows, createSummaryTopStats, createOldModifiedRows, buildSummaryTable, createInfoRows, createSummaryLargeRows, createAnalyticsStats } from './summaryFunctions';

export function createBatchSummary ( batchData: IBatchData ) : React.ReactElement {

  let partialFlag = batchData.isSignificant === true ? '' : '*';

  let tableRows = [];

  tableRows = createSummaryTopStats( tableRows, batchData.summary, batchData, partialFlag );

  tableRows.push( <tr><td>{ `${ batchData.userInfo.count } ${ partialFlag }`} </td><td>{ `Users who created/modified files` }</td></tr> );

  tableRows = createInfoRows( tableRows, batchData, partialFlag );

  tableRows = createSummaryOldRows( tableRows, batchData.oldCreated.summary, partialFlag );
  
  tableRows = createSummaryLargeRows( tableRows, batchData.oldCreated.summary, partialFlag );

  tableRows.push( <tr><td>{ `<< Breaking News !! >>`} </td><td>{ `only ${ createRatioNote( batchData.large.summary,  '' ) }` }</td></tr> );

  tableRows = createOldModifiedRows( tableRows, batchData.oldModified, partialFlag );
  
  tableRows = createAnalyticsStats( tableRows, batchData );

  return buildSummaryTable( tableRows ) ;

}

