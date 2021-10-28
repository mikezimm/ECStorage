import * as React from 'react';
import styles from '../../ExStorage.module.scss';
import { IExTypesProps } from './IExTypesProps';
import { IExTypesState } from './IExTypesState';
import { IExStorageState, IEXStorageList, IEXStorageBatch, IBatchData, IUserSummary, IFileType } from '../../IExStorageState';
import { escape } from '@microsoft/sp-lodash-subset';

import * as strings from 'ExStorageWebPartStrings';

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

// import { sortObjectArrayByChildNumberKey, sortNumberArray } from '@mikezimm/npmfunctions/dist/Services/Arrays/sorting';
import { sortObjectArrayByChildNumberKey, } from '@mikezimm/npmfunctions/dist/Services/Arrays/sorting';

import { createSlider, createChoiceSlider } from '../../fields/sliderFieldBuilder';

import { getStorageItems, batchSize, createBatchData } from '../../ExFunctions';
import { getSearchedFiles } from '../../ExSearch';

import EsItems from '../items/EsItems';

import Gridcharts from '../GridCharts/Gridcharts';
import { makeTheTimeObject } from '@mikezimm/npmfunctions/dist/Services/Time/timeObject';
import { IGridColumns } from '../GridCharts/IGridchartsProps';

import { createTypeSummary } from '../summary/ExTypeSummary';

import { createTypeRatioNote } from '../summary/summaryFunctions';

import { createItemsHeadingWithTypeIcons } from '../miniComps/components';

//copied pivotStyles from \generic-solution\src\webparts\genericWebpart\components\Contents\Lists\railAddTemplate\component.tsx
const pivotStyles = {
  root: {
    whiteSpace: "normal",
    // marginTop: '1em',
  //   textAlign: "center"
  }};

const pivotHeading1 = 'Type Summary';
const pivotHeading2 = 'Items';
const pivotHeading3 = 'Timeline';

export default class ExTypes extends React.Component<IExTypesProps, IExTypesState> {

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



public constructor(props:IExTypesProps){
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
        selectedType: null,

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

  public render(): React.ReactElement<IExTypesProps> {

    const batches = this.props.batches;
    const typesInfo = this.props.typesInfo;
    const pickedList = this.props.pickedList;
    const pickedWeb = this.props.pickedWeb;


    let types = this.props.typesInfo.types;

    let sortedTypes =  types.map( type => {
      return type;
    });
    sortedTypes = sortObjectArrayByChildNumberKey( sortedTypes, 'dec', 'sizeToCountRatio' );

    const bySize = this.buildTypeTables( this.props.typesInfo.types , 'By Total Size', this.state.rankSlider, this.state.textSearch, 'summary.size' );
    const byCount = this.buildTypeTables( this.props.typesInfo.types , 'By Count', this.state.rankSlider, this.state.textSearch, 'summary.count' );
    const byAvg = this.buildTypeTables( this.props.typesInfo.types , 'By Avg size', this.state.rankSlider, this.state.textSearch, 'avgSize' );
    const byMax = this.buildTypeTables( this.props.typesInfo.types , 'By Max size', this.state.rankSlider, this.state.textSearch, 'maxSize' );
    const byRatio = this.buildTypeTables( this.props.typesInfo.types , null, 3, '', 'sizeToCountRatio' );

    
    let heading = <div  style={{paddingBottom: '20px' }}>
      <h3>These file types have few files taking up the most space</h3>
    { byRatio }<div></div></div>;

    //EsItems
    let component = 
    <div>
      <div className={ styles.inflexWrapCenter}>
        { bySize }
        { byCount }
        { byAvg }
        { byMax }
      </div>
    </div>;

    let sliderTypeCount = this.props.batchData.typesInfo.count < 5 ? null : 
      <div style={{margin: '0px 50px 20px 50px'}}> { createSlider( 'Show Top' , this.state.rankSlider , 3, this.props.typesInfo.count, 1 , this._typeSlider.bind(this), this.state.isLoading, 350) }</div> ;

    let userPanel = null;
    
    if ( this.state.showItems === true ) { 

      let componentPivot = 
      <Pivot
          styles={ pivotStyles }
          linkFormat={PivotLinkFormat.links}
          linkSize={PivotLinkSize.normal}
          defaultSelectedKey={ pivotHeading2 }
          // onLinkClick={this._selectedListDefIndex.bind(this)}
      > 
        <PivotItem headerText={ pivotHeading1 } ariaLabel={pivotHeading1} title={pivotHeading1} itemKey={ pivotHeading1 } keytipProps={ { content: 'Hello', keySequences: ['a','b','c'] } }>
          { createTypeSummary( this.state.selectedType, this.props.batchData ) }
        </PivotItem>

        <PivotItem headerText={ pivotHeading2 } ariaLabel={pivotHeading2} title={pivotHeading2} itemKey={ pivotHeading2 } keytipProps={ { content: 'Hello', keySequences: ['a','b','c'] } }>
          <EsItems 
            pickedWeb  = { this.props.pickedWeb }
            pickedList = { this.props.pickedList }
            theSite = {null }

            items = { this.state.items }
            itemsAreDups = { false }
            itemsAreFolders = { false }
            duplicateInfo = { null }
            heading = { ` of type: ${this.state.items[0].docIcon} ${this.props.heading}` }
            // batches = { batches }
            icons = { [] }

            showHeading = { false } // false because we are putting a heading above the pivot items

            dataOptions = { this.props.dataOptions }
            uiOptions = { this.props.uiOptions }
            sharedItems = { [] }

            itemType = { 'Items' }
            >
          </EsItems>
        </PivotItem>

        <PivotItem headerText={ pivotHeading3 } ariaLabel={pivotHeading3} title={pivotHeading3} itemKey={ pivotHeading3 } keytipProps={ { content: 'Hello', keySequences: ['a','b','c'] } }>
          <Gridcharts
            items = { this.state.items }

            // 0 - Context
            pageContext = { this.props.pageContext }
            wpContext = { this.props.wpContext }
            tenant = { this.props.tenant }
            urlVars = { null }
            today = { makeTheTimeObject('')}

            // 2 - Source and destination list information
            parentListWeb = { this.props.pickedWeb.url }
            parentListTitle = { this.props.pickedList.Title }
            parentListURL = { null}

            esItemsHeading = { ``}

            pickedWeb  = { this.props.pickedWeb }
            pickedList = { this.props.pickedList }

            listName = { null}

            columns = { this.props.columns }

            enableSearch = { true }

            scaleMethod = { 'blink' }

            gridStyles = { this.props.gridStyles }

            //Size courtesy of https://www.netwoven.com/2018/11/13/resizing-of-spfx-react-web-parts-in-different-scenarios/
            WebpartElement = { this.props.WebpartElement }

            // 9 - Other web part options
            WebpartHeight = {this.props.WebpartHeight }
            WebpartWidth = { this.props.WebpartWidth }

            // 1 - Analytics options  
            useListAnalytics = { false }
            analyticsWeb = { strings.analyticsWeb }
            analyticsList = { strings.analyticsList}

            // 9 - Other web part options 
            webPartScenario = { null } //Choice used to create mutiple versions of the webpart.

            allLoaded = {null}

            performance = { null }

            parentListFieldTitles = {null}

            refreshId = { this.props.refreshId }

            dataOptions = { this.props.dataOptions }
            uiOptions = { this.props.uiOptions }

          ></Gridcharts>
        </PivotItem> 
      </Pivot>;

      userPanel = <div className={ styles.exStorage }><Panel
        isOpen={ this.state.showItems === true ? true : false }
        // this prop makes the panel non-modal
        
        isBlocking={true}
        onDismiss={ this._onCloseItems.bind(this) }
        closeButtonAriaLabel="Close"
        type = { PanelType.large }
        isLightDismiss = { true }
        >
          { createItemsHeadingWithTypeIcons( this.state.items, 'Items', `with file type: ${this.state.selectedType.iconName}`,'', [], null )}
          { componentPivot }
      </Panel></div>;
    }

    //check for high sizeToCount ratios

    return (
      <div className={ styles.exStorage } style={{ marginLeft: '25px'}}>
        {/* <div className={ styles.container }> */}

          { heading }
          <h3>All File types in this library</h3>
          <div className={ styles.inflexWrapCenter}>
            <div> { sliderTypeCount } </div>
            <div> { this.buildSearchBox() } </div>
          </div>
          { component }
          { userPanel }
          { this.state.isLoading ? 
              <div>
                {/* { loadingNote }
                { searchSpinner }
                { myProgress } */}
              </div>
            : null
          } 
        {/* </div> */}
      </div>
    );
  }

  private buildSearchBox() {
    /*https://developer.microsoft.com/en-us/fabric#/controls/web/searchbox*/
    let searchBox =  
    <div className={[styles.searchContainer, styles.padLeft20 ].join(' ')} >
      <SearchBox
        className={styles.searchBox}
        styles={{ root: { maxWidth: 200 } }}
        placeholder="Search"
        onSearch={ this._searchForItems.bind(this) }
        onFocus={ () => console.log('this.state',  this.state) }
        onBlur={ () => console.log('onBlur called') }
        onChange={ this._searchForItems.bind(this) }
      />
      <div className={styles.searchStatus}>
        { `Search all ${ this.props.typesInfo.count } types` }
        { /* 'Searching ' + (this.state.searchType !== 'all' ? this.state.filteredTiles.length : ' all' ) + ' items' */ }
      </div>
    </div>;

    return searchBox;

  }

  public _searchForItems = (item): void => {
    //This sends back the correct pivot category which matches the category on the tile.
    let e: any = event;
    console.log('searchForItems: e',e);
    console.log('searchForItems: item', item);
    console.log('searchForItems: this', this);

    this.setState({ textSearch: item });
  }

  private buildTypeTables( types: IFileType[] , data: string, countToShow: number, textSearch: string, sortKey: 'summary.size' | 'summary.count' | 'avgSize' | 'maxSize' | 'sizeToCountRatio'): any {

    let elements = [];
    let tableTitle = data;
    const typesSorted: IFileType[] = sortObjectArrayByChildNumberKey( types, 'dec', sortKey );

    typesSorted.map( ( type, index ) => {

      if ( index < countToShow || textSearch.length > 0 ) {
        
        let typePercent = '';
        let label: any = '';

        switch (sortKey) {
          case 'summary.size':
            typePercent = ( type.summary.sizeP ).toFixed( 0 );
            label = `${type.type}  [ ${ type.summary.sizeLabel} / ${ typePercent }% ]` ;
            break;
        
          case 'summary.count':
            typePercent = ( type.summary.countP ).toFixed( 0 );
            label = `${type.type}  [ ${ type.summary.count} / ${ typePercent }% ]` ;
            break;

          case 'avgSize':
            label = `${type.type}  [ ${ type.avgSizeLabel } ]` ;
            break;
        
          case 'maxSize':
            label = `${type.type}  [ ${ type.maxSizeLabel } ]` ;
            break;
        
          case 'sizeToCountRatio':
            label = createTypeRatioNote( type,  '' ) ;
            break;
        
          default:
            break;
        }
//                  { <Icon iconName= { type.type} style={{ padding: '0px 4px 0px 10px', }}></Icon> }
        let showType = textSearch.length === 0 || (textSearch.length > 0 && type.type.toLowerCase().indexOf(textSearch.toLowerCase() )  > -1  ) ? true : false;

        let liStyle : React.CSSProperties = showType === true ?
        {
          display: 'flex',
          flexDirection: 'row',
          justifyContent: 'flex-start',
          alignItems: 'center',
        } : { display: 'none' };

        let newElement = null;
        if ( sortKey === 'sizeToCountRatio' ) {
            newElement = <li style= { liStyle } onClick={ this._onClickItems.bind(this)} id={ type.type }>
            <span style={{width: '30px', paddingRight: '10px'}}>{ index + 1 }. </span><span>{ label }</span>
          </li>;
        } else {
          newElement = <li title={ `${label}` } style= { liStyle } onClick={ this._onClickItems.bind(this)} id={ type.type }>
            <span style={{width: '30px', paddingRight: '10px'}}>{ index + 1 }. </span><span>{ label }</span>
          </li>;
        }
        elements.push( newElement ) ;
      }

    });

    let table = <div style={{marginRight: '10px'}}>
      <h3 style={{ textAlign: 'center' }}> { tableTitle }</h3>
      <ul style={{padding: '0 20px'}}>
        { elements }
      </ul>
    </div>;
    return table;

  }
  
  private _typeSlider(newValue: number){
    this.setState({
      rankSlider: newValue,
    });
  }

  private _onClickItems( event ){
    console.log( event );
    console.log( event.currentTarget.id );
    let showThisType = event.currentTarget.id;
    let items = [];
    let selectedType : IFileType = null;
    this.props.typesInfo.types.map( type => {
      if ( type.type === showThisType ) { 
        items = type.items ;
        selectedType = type;
      }
    });
    this.setState({
      showItems: true,
      items: items,
      selectedType: selectedType,
    });
  }

  private _onCloseItems( event ){
    this.setState({
      showItems: false,
    });
  }

}
