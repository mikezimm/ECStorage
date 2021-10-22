import * as React from 'react';
import styles from '../../ExStorage.module.scss';
import { IExVersionsProps } from './IExVersionsProps';
import { IExVersionsState } from './IExVersionsState';
import { IExStorageState, IEXStorageList, IEXStorageBatch, IBatchData, IUserSummary, IFileType, IVersionBucketLabel } from '../../IExStorageState';
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

import { sortObjectArrayByChildNumberKey, sortNumberArray } from '@mikezimm/npmfunctions/dist/Services/Arrays/sorting';

import { createSlider, createChoiceSlider } from '../../fields/sliderFieldBuilder';

import { getStorageItems, batchSize, createBatchData } from '../../ExFunctions';
import { getSearchedFiles } from '../../ExSearch';

import { createVersionsSummary } from '../summary/ExVersionsSummary';

import EsItems from '../items/EsItems';

const pivotStyles = {
  root: {
    whiteSpace: "normal",
    marginTop: '1em',
  //   textAlign: "center"
  }};

const pivotHeading1 = 'Version Summary';
const pivotHeading2 :IVersionBucketLabel = 'IsDraft';
const pivotHeading3 :IVersionBucketLabel = '1.0';
const pivotHeading4 :IVersionBucketLabel = '>1.0';
const pivotHeading5 :IVersionBucketLabel = '>=100';
const pivotHeading6 :IVersionBucketLabel = '>=500';
const pivotHeading7 :IVersionBucketLabel = 'IsMinor';
const pivotHeading8 :IVersionBucketLabel = 'CheckedOut';


export default class ExVersions extends React.Component<IExVersionsProps, IExVersionsState> {

  private currentDate = new Date();
  private currentYear = this.currentDate.getFullYear();
  
/***
 *          .o88b.  .d88b.  d8b   db .d8888. d888888b d8888b. db    db  .o88b. d888888b  .d88b.  d8888b. 
 *         d8P  Y8 .8P  Y8. 888o  88 88'  YP `~~88~~' 88  `8D 88    88 d8P  Y8 `~~88~~' .8P  Y8. 88  `8D 
 *         8P      88    88 88V8o 88 `8bo.      88    88oobY' 88    88 8P         88    88    88 88oobY' 
 *         8b      88    88 88 V8o88   `Y8b.    88    88`8b   88    88 8b         88    88    88 88`8b   
 *         Y8b  d8 `8b  d8' 88  V888 db   8D    88    88 `88. 88b  d88 Y8b  d8    88    `8b  d8' 88 `88. 
 *          `Y88P'  `Y88P'  VP   V8P `8888Y'    YP    88   YD ~Y8888P'  `Y88P'    YP     `Y88P'  88   YD 
 *                                                                                                       
 *                                                                                                       
 */



public constructor(props:IExVersionsProps){
  super(props);

  let currentYear = new Date();
  let currentYearVal = currentYear.getFullYear();

  this.state = {

        isLoaded: true,
        isLoading: false,
        errorMessage: '',

        hasError: false,
      
        showPane: false,

        items: [],
        showItems: false,

        minYear: currentYearVal - 5 ,
        maxYear: currentYearVal + 1 ,

        rankSlider: 5,
        textSearch: '',

        fetchSlider: 0,
        fetchTotal: 0,
        fetchCount: 0,
        showProgress: false,
        fetchPerComp: 100,
        fetchLabel: '',
  
  };
}


public componentDidMount() {

  // this.updateWebInfo( this.state.parentWeb );
}

//        
  /***
 *         d8888b. d888888b d8888b.      db    db d8888b. d8888b.  .d8b.  d888888b d88888b 
 *         88  `8D   `88'   88  `8D      88    88 88  `8D 88  `8D d8' `8b `~~88~~' 88'     
 *         88   88    88    88   88      88    88 88oodD' 88   88 88ooo88    88    88ooooo 
 *         88   88    88    88   88      88    88 88~~~   88   88 88~~~88    88    88~~~~~ 
 *         88  .8D   .88.   88  .8D      88b  d88 88      88  .8D 88   88    88    88.     
 *         Y8888D' Y888888P Y8888D'      ~Y8888P' 88      Y8888D' YP   YP    YP    Y88888P 
 *                                                                                         
 *                                                                                         
 */

  public componentDidUpdate(prevProps){

  }

  public render(): React.ReactElement<IExVersionsProps> {

    let emptyItemsElements = [
      <div style={{ padding: '20px', width: '100%', height: '100px' }}>
        Well I don't see any files in this category yet.  Is that a good thing?
      </div>,
      <div style={{ padding: '20px', }}>
        I'll tell you one thing about the universe, though. The universe is a pretty big place. It's bigger than anything anyone has ever dreamed of before. So if it's just us... seems like an awful waste of space. Right?
        <br/><br/>- Ellie Arroway
      </div>,
      <div style={{ padding: '20px', }}>
        Looks like we have not created any files this big yet :)
        <br/><br/>Hint - The Tabs tell you how many items fall under this category.
      </div>,
    ];

    let componentPivot = 
    <Pivot
        styles={ pivotStyles }
        linkFormat={PivotLinkFormat.links}
        linkSize={PivotLinkSize.normal}
        // onLinkClick={this._selectedListDefIndex.bind(this)}
    > 
      <PivotItem headerText={ pivotHeading1 } ariaLabel={pivotHeading1} title={pivotHeading1} itemKey={ pivotHeading1 } keytipProps={ { content: 'Hello', keySequences: ['a','b','c'] } }>
        { createVersionsSummary( this.props.batchData.versionInfo, this.props.batchData ) }
      </PivotItem>

      
      <PivotItem headerText={ pivotHeading6 } ariaLabel={pivotHeading6} title={pivotHeading6}
        itemKey={ pivotHeading6 } keytipProps={ { content: 'Hello', keySequences: ['a','b','c'] } } itemCount= { this.props.versionInfo.GT500.length }>
        <EsItems 
            pickedWeb  = { this.props.pickedWeb }
            pickedList = { this.props.pickedList }
            theSite = {null }

            items = { this.props.versionInfo.GT500 }
            itemsAreDups = { false }
            itemsAreFolders = { false }
            duplicateInfo = { null }
            heading = { ` with 500 or more versions` }
            // batches = { batches }
            icons = { [ ]}
            emptyItemsElements = { emptyItemsElements }
                          
            dataOptions = { this.props.dataOptions }
            uiOptions = { this.props.uiOptions }

            sharedItems = { [] }
            
            itemType = { 'Items' }
            
          ></EsItems>
      </PivotItem>

      <PivotItem headerText={ pivotHeading5 } ariaLabel={pivotHeading5} title={pivotHeading5}
        itemKey={ pivotHeading5 } keytipProps={ { content: 'Hello', keySequences: ['a','b','c'] } } itemCount= { this.props.versionInfo.GT100.length }>
        <EsItems 
            pickedWeb  = { this.props.pickedWeb }
            pickedList = { this.props.pickedList }
            theSite = {null }

            items = { this.props.versionInfo.GT100 }
            itemsAreDups = { false }
            itemsAreFolders = { false }
            duplicateInfo = { null }
            heading = { ` with 100-500 versions` }
            // batches = { batches }
            icons = { [ ]}
            emptyItemsElements = { emptyItemsElements }
                          
            dataOptions = { this.props.dataOptions }
            uiOptions = { this.props.uiOptions }

            sharedItems = { [] }
            
            itemType = { 'Items' }
            
          ></EsItems>
      </PivotItem>

      <PivotItem headerText={ pivotHeading2 } ariaLabel={pivotHeading2} title={pivotHeading2} 
        itemKey={ pivotHeading2 } keytipProps={ { content: 'Hello', keySequences: ['a','b','c'] } } itemCount= { this.props.versionInfo.draft.length }>
        <EsItems 
            pickedWeb  = { this.props.pickedWeb }
            pickedList = { this.props.pickedList }
            theSite = {null }

            items = { this.props.versionInfo.draft }
            itemsAreDups = { false }
            itemsAreFolders = { false }
            duplicateInfo = { null }
            heading = { ` in draft state.  Version < 1.0` }
            // batches = { batches }
            icons = { [ ]}
            emptyItemsElements = { emptyItemsElements }
              
            dataOptions = { this.props.dataOptions }
            uiOptions = { this.props.uiOptions }

            sharedItems = { [] }
            
            itemType = { 'Items' }

          ></EsItems>
      </PivotItem>


      <PivotItem headerText={ pivotHeading4 } ariaLabel={pivotHeading4} title={pivotHeading4} 
        itemKey={ pivotHeading4 } keytipProps={ { content: 'Hello', keySequences: ['a','b','c'] } } itemCount= { this.props.versionInfo.GT1.length }>
        <EsItems 
            pickedWeb  = { this.props.pickedWeb }
            pickedList = { this.props.pickedList }
            theSite = {null }

            items = { this.props.versionInfo.GT1 }
            itemsAreDups = { false }
            itemsAreFolders = { false }
            duplicateInfo = { null }
            heading = { ` with versions 1.1 to 99.500` }
            // batches = { batches }
            icons = { [ ]}
            emptyItemsElements = { emptyItemsElements }
                          
            dataOptions = { this.props.dataOptions }
            uiOptions = { this.props.uiOptions }

            sharedItems = { [] }

            itemType = { 'Items' }

          ></EsItems>
      </PivotItem> 
     
      <PivotItem headerText={ pivotHeading3 } ariaLabel={pivotHeading3} title={pivotHeading3}
        itemKey={ pivotHeading3 } keytipProps={ { content: 'Hello', keySequences: ['a','b','c'] } } itemCount= { this.props.versionInfo.one.length }>
        <EsItems 
            pickedWeb  = { this.props.pickedWeb }
            pickedList = { this.props.pickedList }
            theSite = {null }

            items = { this.props.versionInfo.one }
            itemsAreDups = { false }
            itemsAreFolders = { false }
            duplicateInfo = { null }
            heading = { ` with only 1 major version` }
            // batches = { batches }
            icons = { [ ]}
            emptyItemsElements = { emptyItemsElements }
                          
            dataOptions = { this.props.dataOptions }
            uiOptions = { this.props.uiOptions }

            sharedItems = { [] }
            
            itemType = { 'Items' }

          ></EsItems>
      </PivotItem> 
     
      <PivotItem headerText={ pivotHeading7 } ariaLabel={pivotHeading7} title={pivotHeading7}
        itemKey={ pivotHeading7 } keytipProps={ { content: 'Hello', keySequences: ['a','b','c'] } } itemCount= { this.props.versionInfo.minor.length }>
        <EsItems 
            pickedWeb  = { this.props.pickedWeb }
            pickedList = { this.props.pickedList }
            theSite = {null }

            items = { this.props.versionInfo.minor }
            itemsAreDups = { false }
            itemsAreFolders = { false }
            duplicateInfo = { null }
            heading = { ` with minor versions > 1.0` }
            // batches = { batches }
            icons = { [ ]}
            emptyItemsElements = { emptyItemsElements }
                          
            dataOptions = { this.props.dataOptions }
            uiOptions = { this.props.uiOptions }

            sharedItems = { [] }
            
            itemType = { 'Items' }

          ></EsItems>
      </PivotItem> 
     
      <PivotItem headerText={ pivotHeading8 } ariaLabel={pivotHeading8} title={pivotHeading8}
        itemKey={ pivotHeading8 } keytipProps={ { content: 'Hello', keySequences: ['a','b','c'] } } itemCount= { this.props.versionInfo.checkedOut.length }>
        <EsItems 
            pickedWeb  = { this.props.pickedWeb }
            pickedList = { this.props.pickedList }
            theSite = {null }

            items = { this.props.versionInfo.checkedOut }
            itemsAreDups = { false }
            itemsAreFolders = { false }
            duplicateInfo = { null }
            heading = { ` currently Checked Out` }
            // batches = { batches }
            icons = { [ ]}
            emptyItemsElements = { emptyItemsElements }
                          
            dataOptions = { this.props.dataOptions }
            uiOptions = { this.props.uiOptions }

            sharedItems = { [] }
            
            itemType = { 'Items' }

          ></EsItems>
      </PivotItem> 

    </Pivot>;

    return (
      <div className={ styles.exStorage } style={{ marginLeft: '25px'}}>
        {/* <div className={ styles.container }> */}
          {/* <div> */}
            {/* <h3>The larger files</h3> */}
            {/* <p> { this.props.typesInfo.typeList.join(', ') }</p> */}
          {/* </div> */}
          { componentPivot }

      </div>
    );
  }

}
