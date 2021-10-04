import * as React from 'react';
import styles from '../../ExStorage.module.scss';
import { IExUserProps } from './IExUserProps';
import { IExUserState } from './IExUserState';
import { IExStorageState, IEXStorageList, IEXStorageBatch, IBatchData, IUserSummary } from '../../IExStorageState';
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
import { getCountLabel } from '@mikezimm/npmfunctions/dist/Services/Math/basicOperations';

import * as strings from 'ExStorageWebPartStrings';

import { createSlider, createChoiceSlider } from '../../fields/sliderFieldBuilder';

import { getStorageItems, batchSize, createBatchData } from '../../ExFunctions';
import { getSearchedFiles } from '../../ExSearch';

import { createUserSummary } from '../summary/ExUserSummary';

import Gridcharts from '../GridCharts/Gridcharts';
import { makeTheTimeObject } from '@mikezimm/npmfunctions/dist/Services/Time/timeObject';
import { IGridColumns } from '../GridCharts/IGridchartsProps';

import ExTypes from '../types/ExTypes';
import ExSize from '../size/ExSize';
import ExAge from '../age/ExAge';
import ExDups from '../dups/ExDups';
import EsItems from '../items/EsItems';

//copied pivotStyles from \generic-solution\src\webparts\genericWebpart\components\Contents\Lists\railAddTemplate\component.tsx
const pivotStyles = {
  root: {
    whiteSpace: "normal",
    marginTop: '1em',
  //   textAlign: "center"
  }};

const pivotHeading1 = 'User Summary';
const pivotHeading2 = 'Types';
const pivotHeading3 = 'Users';
const pivotHeading4 = 'Size';
const pivotHeading5 = 'Age';
const pivotHeading6 = 'You';
const pivotHeading7 = 'Perms';
const pivotHeading8 = 'Dups';
const pivotHeading9 = 'Folders';
const pivotHeading10 = 'All Files';
const pivotHeading11 = 'Timeline';



export default class ExUser extends React.Component<IExUserProps, IExUserState> {

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



public constructor(props:IExUserProps){
  super(props);

  let currentYear = new Date();
  let currentYearVal = currentYear.getFullYear();

  this.state = {

        isLoaded: false,
        isLoading: true,
        errorMessage: '',

        hasError: false,
      
        showPane: false,

        items: [],

        minYear: currentYearVal - 5 ,
        maxYear: currentYearVal + 1 ,
        yearSlider: currentYearVal,

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

  public render(): React.ReactElement<IExUserProps> {

    const batches = this.props.batches;
    const userSummary = this.props.userSummary;
    const pickedList = this.props.pickedList;
    const pickedWeb = this.props.pickedWeb;

    let componentHeading = <div className={ [styles.flexWrapStart, styles.margRight40Divs ].join(' ') }>
      <div className={ [styles.inflexNoWrapStart, styles.margRight15Divs ].join(' ') }>
        <div style={{whiteSpace: 'nowrap'}} >[ { this.props.userSummary.userId } ]</div>
        <div style={{whiteSpace: 'nowrap'}}>{ this.props.userSummary.userTitle }</div>
      </div>

      <div className={ [styles.inflexNoWrapStart, styles.margRight15Divs ].join(' ') }>
        <div>Created</div>
        <div style={{whiteSpace: 'nowrap'}}>{ getCountLabel(userSummary.createCount) } / { userSummary.summary.countP.toFixed(1) }% files</div>
        <div style={{whiteSpace: 'nowrap'}}>{ userSummary.createTotalSizeLabel } / { userSummary.summary.sizeP.toFixed(1) }%</div>
      </div>

      <div className={ [styles.inflexNoWrapStart, styles.margRight15Divs ].join(' ') }>
        <div>Modified</div>
        <div style={{whiteSpace: 'nowrap'}}>{ getCountLabel(userSummary.modifyCount) } files</div>
        <div style={{whiteSpace: 'nowrap'}}>{ userSummary.modifyTotalSizeLabel }</div>
      </div>

    </div>;
    // let timeComment = null;
    // let etaMinutes = pickedList && this.state.fetchSlider > 0 ? (  this.state.fetchSlider * 7 / ( 1000 * 60 ) ).toFixed( 1 ) : 0;
    // if ( this.state.isLoading ) {
    //   timeComment = etaMinutes ;
    // }

    // let loadingNote = this.state.isLoading ? <div>
    //   Please do not interupt the process which could take { etaMinutes } minutes.
    // </div> : null;
    // let searchSpinner = this.state.isLoading ? 
    //     <Spinner size={SpinnerSize.large} label={` fetching ${pickedList ? this.state.fetchSlider : 'TBD'} items...`} />
    //  : null ;

    // let myProgress = 1 === 1 ? <ProgressIndicator 
    // label={ this.state.fetchLabel } 
    // description={ '' } 
    // percentComplete={ this.state.fetchPerComp } 
    // progressHidden={ !this.state.showProgress }/> : null;


    // let fetchButton = <PrimaryButton text={ 'Begin'} onClick={this.fetchStoredItemsClick.bind(this)} allowDisabledFocus disabled={ this.state.isLoading } />;

    // let currentYear = new Date();
    // let currentYearVal = currentYear.getFullYear();

    // let sliderYearItself = !pickedList ? null : 
    //   <div style={{margin: '0 50px'}}> { createSlider( null , this.state.yearSlider , this.state.minYear, this.state.maxYear, 1 , this._updateMaxYear.bind(this), this.state.isLoading, 350) }</div> ;

    // let sliderYearComponent = !pickedList ? null : <div style={{display:'inline-flex', alignItems: 'center', justifyContent: 'center', flexWrap: 'wrap', marginTop: '20px' }}>
    //   <span style={{ fontSize: 'larger', fontWeight: 'bolder', minWidth: '300px' }}> { `Ignore files Created > ${ this.state.yearSlider }` } </span>
    //   { sliderYearItself }
    // </div>;

    // let sliderCountItself = !pickedList ? null : 
    //   <div style={{margin: '0 50px'}}> { createSlider( null , this.state.fetchSlider , 0, pickedList.ItemCount, batchSize , this._updateMaxFetch.bind(this), this.state.isLoading, 350) }</div> ;

    // let sliderCountComponent = !pickedList ? null : <div style={{display:'inline-flex', alignItems: 'center', justifyContent: 'center', flexWrap: 'wrap', marginTop: '20px' }}>
    //   <span style={{ fontSize: 'larger', fontWeight: 'bolder', minWidth: '300px' }}> { `Fetch up to ${pickedList.ItemCount } Files` } </span>
    //   { sliderCountItself }
    //   <span style={{marginRight: '50px'}}> { `Plan for about ${etaMinutes} minutes` } </span>
    //   { fetchButton }
    // </div>;
    let heading = `, from ${this.props.userSummary.userTitle}`;  //          heading = { heading }
    let typesPivotContent = <div>
      <ExTypes 
  
          //Size courtesy of https://www.netwoven.com/2018/11/13/resizing-of-spfx-react-web-parts-in-different-scenarios/
          WebpartHeight = { this.props.WebpartHeight }
          WebpartWidth = { this.props.WebpartWidth }
      
          pickedWeb  = { this.props.pickedWeb }
          pickedList = { this.props.pickedList }
          theSite = {null }

          typesInfo = { userSummary.typesInfo }
          batches = { batches }
          batchData = { this.props.batchData }
          heading = { heading }

          dataOptions = { this.props.dataOptions }
          uiOptions = { this.props.uiOptions }
      >
      </ExTypes>
      <ReactJson src={ userSummary.typesInfo.types } name={ 'Types' } collapsed={ true } displayDataTypes={ true } displayObjectSize={ true } enableClipboard={ true } style={{ padding: '20px 0px' }}/></div>;

    let usersPivotContent = null;

    let sizePivotContent = <div>
      <ExSize 
        //Size courtesy of https://www.netwoven.com/2018/11/13/resizing-of-spfx-react-web-parts-in-different-scenarios/
        WebpartHeight = { this.props.WebpartHeight }
        WebpartWidth = { this.props.WebpartWidth }

        pickedWeb  = { this.props.pickedWeb }
        pickedList = { this.props.pickedList }
        theSite = {null }

        batchData = { this.props.batchData }

        largeData = { userSummary.large }
        
        heading = { heading }

        dataOptions = { this.props.dataOptions }
        uiOptions = { this.props.uiOptions }

      >
      </ExSize></div>;

      let agePivotContent = <div>
        <ExAge
          //Size courtesy of https://www.netwoven.com/2018/11/13/resizing-of-spfx-react-web-parts-in-different-scenarios/
          WebpartHeight = { this.props.WebpartHeight }
          WebpartWidth = { this.props.WebpartWidth }

          pickedWeb  = { this.props.pickedWeb }
          pickedList = { this.props.pickedList }
          theSite = {null }

          batchData = { this.props.batchData }

          oldFiles = { userSummary.oldCreated }
                  
          heading = { heading }

          dataOptions = { this.props.dataOptions }
          uiOptions = { this.props.uiOptions }
        >
        </ExAge></div>;


    let permsPivotContent = <div><div>
      <h3>Summary of files with broken permissions</h3>
      </div>
        <ReactJson src={ userSummary.uniqueInfo.uniqueRolls} name={ 'Broken Permissions' } collapsed={ true } displayDataTypes={ true } displayObjectSize={ true } enableClipboard={ true } style={{ padding: '20px 0px' }}/>
      </div>;


    let dupsPivotContent = <div>
      <ExDups
        //Size courtesy of https://www.netwoven.com/2018/11/13/resizing-of-spfx-react-web-parts-in-different-scenarios/
        WebpartHeight = { this.props.WebpartHeight }
        WebpartWidth = { this.props.WebpartWidth }

        pickedWeb  = { this.props.pickedWeb }
        pickedList = { this.props.pickedList }
        theSite = {null }

        heading = { '' }

        batchData = { this.props.batchData }

        duplicateInfo = { userSummary.duplicateInfo }
                                
        dataOptions = { this.props.dataOptions }
        uiOptions = { this.props.uiOptions }

      >
      </ExDups></div>;

    let folderPivotContent = <div><div>
      <h3>Summary of Folders</h3>
      </div>
        <ReactJson src={ userSummary.folderInfo.folders} name={ 'Folders' } collapsed={ true } displayDataTypes={ true } displayObjectSize={ true } enableClipboard={ true } style={{ padding: '20px 0px' }}/>
      </div>;

    let itemsContent = <EsItems 

      pickedWeb  = { this.props.pickedWeb }
      pickedList = { this.props.pickedList }
      theSite = {null }

      items = { this.props.userSummary.items }
      itemsAreDups = { false }
      duplicateInfo = { null }
      heading = { `From user: ${ this.props.userSummary.userTitle }` }
      // batches = { batches }
      icons = { [] }

      dataOptions = { this.props.dataOptions }
      uiOptions = { this.props.uiOptions }

      >
    </EsItems>;
          
    let gridPivotContent = !this.props.isLoaded || this.props.userSummary.summary.count === 0 ? null : 
    <Gridcharts

      items = { this.props.userSummary.items }

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

    ></Gridcharts>;

    let componentPivot = 
    <Pivot
        styles={ pivotStyles }
        linkFormat={PivotLinkFormat.links}
        linkSize={PivotLinkSize.normal}
        // onLinkClick={this._selectedListDefIndex.bind(this)}
    > 
      <PivotItem headerText={ pivotHeading1 } ariaLabel={pivotHeading1} title={pivotHeading1} itemKey={ pivotHeading1 } keytipProps={ { content: 'Hello', keySequences: ['a','b','c'] } }>
        { createUserSummary( this.props.userSummary, this.props.batchData ) }
      </PivotItem>

      <PivotItem headerText={ pivotHeading2 } ariaLabel={pivotHeading2} title={pivotHeading2} itemKey={ pivotHeading2 } keytipProps={ { content: 'Hello', keySequences: ['a','b','c'] } }>
        { typesPivotContent }
      </PivotItem>

      <PivotItem headerText={ pivotHeading4 } ariaLabel={pivotHeading4} title={pivotHeading4} itemKey={ pivotHeading4 } keytipProps={ { content: 'Hello', keySequences: ['a','b','c'] } }>
        { sizePivotContent }
      </PivotItem> 
      
      <PivotItem headerText={ pivotHeading5 } ariaLabel={pivotHeading5} title={pivotHeading5} itemKey={ pivotHeading5 } keytipProps={ { content: 'Hello', keySequences: ['a','b','c'] } }>
        { agePivotContent }
      </PivotItem>

      <PivotItem headerText={ pivotHeading7 } ariaLabel={pivotHeading7} title={pivotHeading7} itemKey={ pivotHeading7 } keytipProps={ { content: 'Hello', keySequences: ['a','b','c'] } }>
        { permsPivotContent }
      </PivotItem>

      <PivotItem headerText={ pivotHeading8 } ariaLabel={pivotHeading8} title={pivotHeading8} itemKey={ pivotHeading8 } keytipProps={ { content: 'Hello', keySequences: ['a','b','c'] } }>
        { dupsPivotContent }
      </PivotItem>

      <PivotItem headerText={ pivotHeading9 } ariaLabel={pivotHeading9} title={pivotHeading9} itemKey={ pivotHeading9 } keytipProps={ { content: 'Hello', keySequences: ['a','b','c'] } }>
        { folderPivotContent }
      </PivotItem>

      <PivotItem headerText={ pivotHeading10 } ariaLabel={pivotHeading10} title={pivotHeading10} itemKey={ pivotHeading10 } keytipProps={ { content: 'Hello', keySequences: ['a','b','c'] } }>
        { itemsContent }
      </PivotItem>

      <PivotItem headerText={ pivotHeading11 } ariaLabel={pivotHeading11} title={pivotHeading11} itemKey={ pivotHeading11 } keytipProps={ { content: 'Hello', keySequences: ['a','b','c'] } }>
        { gridPivotContent }
      </PivotItem>
    </Pivot>;

    return (
      <div className={ styles.exStorage }>
        <div className={ styles.container } style={{ paddingTop: '1em !important'}}>

          {/* <span className={ styles.title }>Welcome to SharePoint!</span> */}
          {/* <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p> */}
          {/* <p className={ styles.description }>{escape(this.props.parentWeb)}</p> */}
          {/* <div>{ this.state.currentUser ? this.state.currentUser.Title : null }</div> */}

          {/* { sliderYearComponent }
          { sliderCountComponent } */}
          { componentHeading }
          { this.state.isLoading ? 
              <div>
                {/* { loadingNote }
                { searchSpinner }
                { myProgress } */}
              </div>
            : null
          } 

          { componentPivot }
          <div style={{ overflowY: 'auto' }}>
              {/* <ReactJson src={ this.state.currentUser } collapsed={ true } displayDataTypes={ true } displayObjectSize={ true } enableClipboard={ true } style={{ padding: '20px 0px' }}/>
              <ReactJson src={ this.state.pickedWeb } collapsed={ true } displayDataTypes={ true } displayObjectSize={ true } enableClipboard={ true } style={{ padding: '20px 0px' }}/>
              <ReactJson src={ pickedList } collapsed={ true } displayDataTypes={ true } displayObjectSize={ true } enableClipboard={ true } style={{ padding: '20px 0px' }}/> */}
              <ReactJson src={ batches } name={ 'All items' } collapsed={ true } displayDataTypes={ true } displayObjectSize={ true } enableClipboard={ true } style={{ padding: '20px 0px' }}/>

          </div>
        </div>
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
        { `Add search label here for ${this.props.currentUser.Title}` }
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

  private buildUserTables( indexs: number[], users: IUserSummary[] , data: string, countToShow: number, userSearch: string ): any {

    let elements = [];
    let tableTitle = data;

    indexs.map( ( allUserIndex, index ) => {

      if ( index < countToShow || userSearch.length > 0 ) {
        let user = users[ allUserIndex ];
        let label = '' ;
        let createTotalSize = user.createTotalSize > 1e9 ? `${ (user.createTotalSize / 1e9).toFixed(1) } GB` : `${ (user.createTotalSize / 1e6).toFixed(1) } MB`;
        let modifyTotalSize = user.modifyTotalSize > 1e9 ? `${ (user.modifyTotalSize / 1e9).toFixed(1) } GB` : `${ (user.modifyTotalSize / 1e6).toFixed(1) } MB`;
        let createPercent = ( user.summary.sizeP ).toFixed( 0 );
        let countPercent = ( user.summary.countP ).toFixed( 0 );

        switch (data) {
          case 'createSizeRank':
            label = `${user.userTitle}  [ ${ createTotalSize } / ${ createPercent }% ]` ;
            break;
          case 'createCountRank':
            label = `${user.userTitle}  [ ${user.createCount} / ${ countPercent }% ]` ;
            break;
          case 'modifySizeRank':
            label = `${user.userTitle}  [ ${ modifyTotalSize } ]` ;
            break;
          case 'modifyCountRank':
            label = `${user.userTitle}  [ ${user.modifyCount} ]` ;
            break;
  
          default:
            break;
        }
        
        let title = `( #${ allUserIndex } Id: ${user.userId} ) ${user.userTitle} created ${ createTotalSize }, modified ${modifyTotalSize}` ;

        let showUser = userSearch.length === 0 || (userSearch.length > 0 && user.userTitle.toLowerCase().indexOf(userSearch.toLowerCase() )  > -1  ) ? true : false;

        let liStyle : React.CSSProperties = showUser === true ?
        {
          display: 'flex',
          flexDirection: 'row',
          justifyContent: 'flex-start',
          alignItems: 'center',
        } : { display: 'none' };

        elements.push(<li title={ title} style= { liStyle }>
          <span style={{width: '30px', paddingRight: '10px'}}>{ index + 1 }. </span><span>{ label }</span>
        </li>);

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
  
  private _updateMaxYear(newValue: number){
    this.setState({
      yearSlider: newValue,
    });
  }

  private _updateMaxFetch(newValue: number){
    this.setState({
      fetchSlider: newValue,
    });
  }

  private _updateRankShow(newValue: number){
    this.setState({
      rankSlider: newValue,
    });
  }

}
