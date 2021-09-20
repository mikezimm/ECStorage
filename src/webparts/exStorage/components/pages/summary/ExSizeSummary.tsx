import * as React from 'react';
import styles from '../../ExStorage.module.scss';

import { IExStorageState, IEXStorageList, IEXStorageBatch, IBatchData, IUserSummary, IFileType, ILargeFiles } from '../../IExStorageState';
import { escape } from '@microsoft/sp-lodash-subset';


import { sp, Views, IViews, ISite } from "@pnp/sp/presets/all";
import { Web, IList, Site } from "@pnp/sp/presets/all";

import { ProgressIndicator } from 'office-ui-fabric-react/lib/ProgressIndicator';

import {
  Spinner,
  SpinnerSize,
  FloatingPeoplePicker,
  // MessageBar,
  // MessageBarType,
  // SearchBox,
  // Icon,
  // Label,
  // Pivot,
  // PivotItem,
  // IPivotItemProps,
  // PivotLinkFormat,
  // PivotLinkSize,
  // Dropdown,
  // IDropdownOption
} from "office-ui-fabric-react";
import { Icon  } from 'office-ui-fabric-react/lib/Icon';
import { DefaultButton, PrimaryButton, CompoundButton, Stack, IStackTokens, elementContains } from 'office-ui-fabric-react';
import { SearchBox } from 'office-ui-fabric-react/lib/SearchBox';

import { Panel, IPanelProps, IPanelStyleProps, IPanelStyles, PanelType } from 'office-ui-fabric-react/lib/Panel';

import { Pivot, PivotItem, IPivotItemProps, PivotLinkFormat, PivotLinkSize,} from 'office-ui-fabric-react/lib/Pivot';
import { Dropdown, DropdownMenuItemType, IDropdownStyles, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { MessageBar, MessageBarType,  } from 'office-ui-fabric-react/lib/MessageBar';

import ReactJson from "react-json-view";

import { IPickedWebBasic, IPickedList, }  from '@mikezimm/npmfunctions/dist/Lists/IListInterfaces';
import { IUser } from '@mikezimm/npmfunctions/dist/Services/Users/IUserInterfaces';

import { getSiteInfo, getWebInfoIncludingUnique } from '@mikezimm/npmfunctions/dist/Services/Sites/getSiteInfo';
import { cleanURL } from '@mikezimm/npmfunctions/dist/Services/Strings/urlServices';
import { getHelpfullErrorV2 } from '@mikezimm/npmfunctions/dist/Services/Logging/ErrorHandler';

import { sortObjectArrayByNumberKey, sortNumberArray } from '@mikezimm/npmfunctions/dist/Services/Arrays/sorting';

import { createSlider, createChoiceSlider } from '../../fields/sliderFieldBuilder';

import { getStorageItems, batchSize, createBatchData, getSizeLabel } from '../../ExFunctions';
import { getSearchedFiles } from '../../ExSearch';

import EsItems from '../items/EsItems';


export function createSizeSummary ( large: ILargeFiles, batchData: IBatchData ) : React.ReactElement {
  let fullLoad = large.summary.count === batchData.totalCount ? ' all' : ' ONLY';

  let loadPercent = batchData.totalCount !== 0 ? (( large.summary.count / batchData.totalCount ) * 100) : 0;
  let loadPercentLabel = loadPercent.toFixed(1);
  let partialFlag = loadPercent === 100 ? '' : '*';

  let mainHeading = `Showing results for${fullLoad} ${ large.summary.count } of ${ batchData.totalCount }`;
  let secondHeading = `This represents${fullLoad} ${ loadPercentLabel } of the files in this library.`;
  let tableRows = [];

  tableRows.push( <tr><td>{ `${ large.summary.count } of ${ batchData.totalCount }`} </td><td>{ `Showing results for this many files in the library` }</td></tr> );
  tableRows.push( <tr><td>{ `or ${ loadPercentLabel }%`} </td><td>{ `% of all the files available` }</td></tr> );
  if ( loadPercent !== 100 ) {
    tableRows.push( <tr><td>{ partialFlag } </td><td>{ `Loading only part of the files may provide mis-leading results.` }</td></tr> );
    tableRows.push( <tr><td>{ null } </td><td>{ `For a complete picture, slide the Fetch counter all the way to the right and press Begin button` }</td></tr> );
  }
  tableRows.push( <tr><td>{ `${ large.summary.sizeLabel } ${ partialFlag }`} </td><td>{ `Total size of all files larger than 100MB` }</td></tr> );

  let GT100M = large.summary.count;
  let GT100SizeLabel = getSizeLabel(large.summary.size);

  tableRows.push( <tr><td>{ `${ GT100M } or ${ GT100SizeLabel } ${ partialFlag }`} </td><td>{ `Files larger than 100MB ` }</td></tr> );
  tableRows.push( <tr><td>{ `<< Breaking News !! >>`} </td><td>{ `only ${ large.summary.countP.toFixed(4) }% of all these files account for ${ large.summary.sizeP.toFixed(4) }% of your storage` }</td></tr> );

  let summaryTable = <table className={ styles.summaryTable }>
    { tableRows }
  </table>;

  const totalsInfo = <div className={ styles.flexWrapStart }>
    <div>{ mainHeading }</div>
    <div>{ secondHeading }</div>

  </div>;
  return <div style={{paddingTop: '20px' }}>
    { summaryTable }
  </div>;

}
