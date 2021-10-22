import * as React from 'react';
import styles from '../../ExStorage.module.scss';

import { IExStorageState, IEXStorageList, IEXStorageBatch, IBatchData, IUserSummary, IFileType, ILargeFiles, IOldFiles } from '../../IExStorageState';
import { escape } from '@microsoft/sp-lodash-subset';

import {

} from "office-ui-fabric-react";

import { sortObjectArrayByChildNumberKey, sortNumberArray } from '@mikezimm/npmfunctions/dist/Services/Arrays/sorting';

import { getSizeLabel, getCommaSepLabel } from '@mikezimm/npmfunctions/dist/Services/Math/basicOperations';

import { createRatioNote, createSummaryRangeRows, createSummaryOldRows, createSummaryTopStats, createOldModifiedRows, buildSummaryTable, createInfoRows, createSummaryLargeRows, createSummaryTypeRows } from './summaryFunctions';

export function createTypeSummary ( fileType: IFileType, batchData: IBatchData ) : React.ReactElement {

  let partialFlag = batchData.isSignificant === true ? '' : '*';

  let tableRows = [];

  tableRows = createSummaryTopStats( tableRows, fileType.summary, batchData, partialFlag );

  tableRows = createSummaryTypeRows( tableRows, fileType, partialFlag );

  tableRows = createSummaryRangeRows( tableRows, fileType.summary );

  tableRows.push( <tr><td>{ `<< Breaking News !! >>`} </td><td>{ createRatioNote( fileType.summary,  '' ) }</td></tr> );

  let tableStyle = { marginTop: '1.4em'};

  return buildSummaryTable( tableRows, styles.exStorage, null ) ;

}
