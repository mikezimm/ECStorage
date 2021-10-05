import * as React from 'react';
import styles from '../../ExStorage.module.scss';

import { IExStorageState, IEXStorageList, IEXStorageBatch, IBatchData, IUserSummary, IFileType, ILargeFiles, IOldFiles } from '../../IExStorageState';
import { escape } from '@microsoft/sp-lodash-subset';

import {

} from "office-ui-fabric-react";

import { sortObjectArrayByChildNumberKey, sortNumberArray } from '@mikezimm/npmfunctions/dist/Services/Arrays/sorting';

import { getStorageItems, batchSize, createBatchData, } from '../../ExFunctions';
import { getSizeLabel, getCommaSepLabel } from '@mikezimm/npmfunctions/dist/Services/Math/basicOperations'; 

import { createRatioNote, createSummaryRangeRows, createSummaryOldRows, createSummaryTopStats, createOldModifiedRows, buildSummaryTable, createInfoRows, createSummaryLargeRows } from './summaryFunctions';


export function createSizeSummary ( large: ILargeFiles, batchData: IBatchData ) : React.ReactElement {

  let partialFlag = batchData.isSignificant === true ? '' : '*';

  let tableRows = [];

  tableRows = createSummaryTopStats( tableRows, large.summary, batchData, partialFlag );

  tableRows = createSummaryOldRows( tableRows, large.summary, partialFlag );
  
  tableRows = createSummaryRangeRows( tableRows, large.summary );

  tableRows.push( <tr><td>{ `<< Breaking News !! >>`} </td><td>{ createRatioNote( large.summary,  '' ) }</td></tr> );

  return buildSummaryTable( tableRows ) ;

}
