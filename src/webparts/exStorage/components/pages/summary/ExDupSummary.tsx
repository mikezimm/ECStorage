import * as React from 'react';
import styles from '../../ExStorage.module.scss';

import { IExStorageState, IEXStorageList, IEXStorageBatch, IBatchData, IUserSummary, IFileType, ILargeFiles, IOldFiles, IDuplicateInfo } from '../../IExStorageState';
import { escape } from '@microsoft/sp-lodash-subset';

import {

} from "office-ui-fabric-react";

import { sortObjectArrayByChildNumberKey, sortNumberArray } from '@mikezimm/npmfunctions/dist/Services/Arrays/sorting';
import { getSizeLabel, getCommaSepLabel } from '@mikezimm/npmfunctions/dist/Services/Math/basicOperations'; 

import { getStorageItems, batchSize, createBatchData, } from '../../ExFunctions';

import { createRatioNote, createSummaryRangeRows, createSummaryOldRows, createSummaryTopStats } from './summaryFunctions';


export function createDupSummary ( dups: IDuplicateInfo, batchData: IBatchData ) : React.ReactElement {

  let partialFlag = batchData.isSignificant === true ? '' : '*';

  let tableRows = [];

  tableRows = createSummaryTopStats( tableRows, dups.summary, batchData, partialFlag );

  tableRows = createSummaryOldRows( tableRows, dups.summary, partialFlag );

  tableRows = createSummaryRangeRows( tableRows, dups.summary );

  tableRows.push( <tr><td>{ `<< Breaking News !! >>`} </td><td>{ createRatioNote( dups.summary,  '' ) }</td></tr> );

  let summaryTable = <table className={ styles.summaryTable }>
    { tableRows }
  </table>;

  return <div style={{  }}>
    { summaryTable }
  </div>;

}
