
import * as React from 'react';
import styles from '../../ExStorage.module.scss';

import { IExStorageState, IEXStorageList, IEXStorageBatch, IBatchData, IUserSummary, IFileType, IBucketSummary } from '../../IExStorageState';


export function createRatioNote ( summary: IBucketSummary, userLabel: string ) {
  if ( !userLabel || userLabel.length === 0 ) { userLabel = 'all' ; }
  return  `only ${ summary.countP.toPrecision(2) }% of ${ userLabel } files ( ${summary.count} ) account for ${ summary.sizeP.toPrecision(2) }%  ( ${summary.sizeLabel} ) of ${ userLabel } space`;
}

export function createTypeRatioNote ( summary: IFileType, userLabel: string ) {  //sizeToCountRatio
  if ( !userLabel || userLabel.length === 0 ) { userLabel = 'all' ; }
  let text = `${ summary.type }:  only ${ summary.countP.toPrecision(2) }% of ${ userLabel } files ( ${summary.count} )  account for ${ summary.sizeP.toPrecision(2) }% ( ${summary.sizeLabel} ) of ${ userLabel } space`;
  let title = `The size ( ${ summary.sizeLabel }) to count ( ${ summary.count }) ratio is ${ summary.sizeToCountRatio.toPrecision(2) }`;
  return  <span title={ title }>{ text }</span>;
}