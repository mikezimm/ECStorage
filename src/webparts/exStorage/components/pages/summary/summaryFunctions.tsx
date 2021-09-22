
import * as React from 'react';
import styles from '../../ExStorage.module.scss';

import { IExStorageState, IEXStorageList, IEXStorageBatch, IBatchData, IUserSummary, IFileType, IBucketSummary } from '../../IExStorageState';


export function createRatioNote ( summary: IBucketSummary, userLabel: string ) {
  if ( !userLabel || userLabel.length === 0 ) { userLabel = 'all' ; }
  return  `only ${ summary.countP.toFixed(4) }% of ${ userLabel } files ( ${summary.count} ) account for ${ summary.sizeP.toFixed(4) }% ${ userLabel } space`;
}

export function createTypeRatioNote ( summary: IFileType, userLabel: string ) {  //sizeToCountRatio
  if ( !userLabel || userLabel.length === 0 ) { userLabel = 'all' ; }
  let text = `${ summary.type }:  only ${ summary.countP.toFixed(4) }% of ${ userLabel } files ( ${summary.count} )  account for ${ summary.sizeP.toFixed(4) }% ${ userLabel } space`;
  let title = `The size ( ${ summary.sizeLabel }) to count ( ${ summary.count }) ratio is ${ summary.sizeToCountRatio.toFixed(2) }`;
  return  <span title={ title }>{ text }</span>;
}