import * as React from 'react';
import styles from '../../ExStorage.module.scss';

import { IExStorageState, IEXStorageList, IEXStorageBatch, IBatchData, IUserSummary, IFileType, ILargeFiles, IOldFiles } from '../../IExStorageState';
import { escape } from '@microsoft/sp-lodash-subset';

import {

} from "office-ui-fabric-react";

import { sortObjectArrayByChildNumberKey, sortNumberArray } from '@mikezimm/npmfunctions/dist/Services/Arrays/sorting';

import { getSizeLabel, getCommaSepLabel } from '@mikezimm/npmfunctions/dist/Services/Math/basicOperations';

import { createRatioNote, createSummaryRangeRows, createSummaryOldRows, createSummaryTopStats } from './summaryFunctions';

export function createUserSummary ( userSummary: IUserSummary, batchData: IBatchData ) : React.ReactElement {
  // const summary = userSummary.summary;


  let partialFlag = batchData.isSignificant === true ? '' : '*';

  let tableRows = [];

  tableRows = createSummaryTopStats( tableRows, userSummary.summary, batchData, partialFlag );

  tableRows.push( <tr><td>{ `${ getCommaSepLabel(userSummary.typesInfo.count) } ${ partialFlag }`} </td><td>{ `File types found` }</td></tr> );
  tableRows.push( <tr><td>{ `${ getCommaSepLabel(userSummary.duplicateInfo.count) } ${ partialFlag }`} </td><td>{ `Files that have more than one copy in the library` }</td></tr> );
  tableRows.push( <tr><td>{ `${ getCommaSepLabel(userSummary.uniqueInfo.count) } ${ partialFlag }`} </td><td>{ `Folders/files with Unique Permissions` }</td></tr> );
  // tableRows.push( <tr><td>{ `${ userSummary.count } ${ partialFlag }`} </td><td>{ `Users who created/modified files` }</td></tr> );

  tableRows = createSummaryOldRows( tableRows, userSummary.summary, partialFlag );

  tableRows = createSummaryRangeRows( tableRows, userSummary.summary );

  let userLabel = userSummary.userId === batchData.userInfo.currentUser.userId ? 'your' : 'this user\'s';
  tableRows.push( <tr><td>{ `<< Breaking News !! >>`} </td><td>{ createRatioNote( userSummary.large.summary,  '' )  }</td></tr> );

  let Age3YrCount = userSummary.oldModified.Age3Yr.length;
  Age3YrCount += userSummary.oldModified.Age4Yr.length;
  Age3YrCount += userSummary.oldModified.Age5Yr.length;

  tableRows.push( <tr><td>{ `${ Age3YrCount } ${ partialFlag }`} </td><td>{ `Files last modified more than a couple years ago` }</td></tr> );

  let summaryTable = <table className={ styles.summaryTable }>
    { tableRows }
  </table>;

  return <div style={{  }}>
    { summaryTable }
  </div>;

}