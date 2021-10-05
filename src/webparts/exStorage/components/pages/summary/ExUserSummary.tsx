import * as React from 'react';
import styles from '../../ExStorage.module.scss';

import { IExStorageState, IEXStorageList, IEXStorageBatch, IBatchData, IUserSummary, IFileType, ILargeFiles, IOldFiles } from '../../IExStorageState';
import { escape } from '@microsoft/sp-lodash-subset';

import {

} from "office-ui-fabric-react";

import { sortObjectArrayByChildNumberKey, sortNumberArray } from '@mikezimm/npmfunctions/dist/Services/Arrays/sorting';

import { getSizeLabel, getCommaSepLabel } from '@mikezimm/npmfunctions/dist/Services/Math/basicOperations';

import { createRatioNote, createSummaryRangeRows, createSummaryOldRows, createSummaryTopStats, createOldModifiedRows, buildSummaryTable, createInfoRows, createSummaryLargeRows } from './summaryFunctions';

export function createUserSummary ( userSummary: IUserSummary, batchData: IBatchData ) : React.ReactElement {
  // const summary = userSummary.summary;


  let partialFlag = batchData.isSignificant === true ? '' : '*';

  let tableRows = [];

  tableRows = createSummaryTopStats( tableRows, userSummary.summary, batchData, partialFlag );

  tableRows = createInfoRows( tableRows, batchData, partialFlag );

  tableRows = createSummaryOldRows( tableRows, userSummary.summary, partialFlag );

  tableRows = createSummaryLargeRows( tableRows, userSummary.summary, partialFlag );

  tableRows = createSummaryRangeRows( tableRows, userSummary.summary );

  let userLabel = userSummary.userId === batchData.userInfo.currentUser.userId ? 'your' : 'this user\'s';

  tableRows.push( <tr><td>{ `<< Breaking News !! >>`} </td><td>{ createRatioNote( userSummary.large.summary,  '' )  }</td></tr> );

  tableRows = createOldModifiedRows( tableRows, userSummary.oldModified, partialFlag );

  return buildSummaryTable( tableRows ) ;

}