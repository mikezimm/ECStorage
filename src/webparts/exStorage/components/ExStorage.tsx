import * as React from 'react';
import styles from './ExStorage.module.scss';
import { IExStorageProps } from './IExStorageProps';
import { IExStorageState, IEXStorageList, IEXStorageBatch, IBatchData, IUserSummary } from './IExStorageState';
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

import { createSlider, createChoiceSlider } from './fields/sliderFieldBuilder';

import { getStorageItems, batchSize, createBatchData } from './ExFunctions';
import { getSearchedFiles } from './ExSearch';

import ExUser from './pages/user/ExUser';
import ExTypes from './pages/types/ExTypes';

//copied pivotStyles from \generic-solution\src\webparts\genericWebpart\components\Contents\Lists\railAddTemplate\component.tsx
const pivotStyles = {
  root: {
    whiteSpace: "normal",
    marginTop: '30px',
  //   textAlign: "center"
  }};

const pivotHeading1 = 'Summary';
const pivotHeading2 = 'Types';
const pivotHeading3 = 'Users';
const pivotHeading4 = 'Size';
const pivotHeading5 = 'Age';
const pivotHeading6 = 'You';
const pivotHeading7 = 'Perms';
const pivotHeading8 = 'Dups';
const pivotHeading9 = 'Folders';



export default class ExStorage extends React.Component<IExStorageProps, IExStorageState> {

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



public constructor(props:IExStorageProps){
  super(props);

  let parentWeb = cleanURL(this.props.parentWeb);

  let currentYear = new Date();
  let currentYearVal = currentYear.getFullYear();

  this.state = {

        pickedList : null,
        pickedWeb : this.props.pickedWeb,
        isLoaded: false,
        isLoading: true,
        errorMessage: '',
        stateError: [],
        hasError: false,
      
        showPane: false,
        showUser: -1,
        currentUser: null,
        isCurrentWeb: null,

        parentWeb: parentWeb,
        listTitle: this.props.listTitle,

        allowRailsOff: false,

        theSite: null,

        batches: [],
        items: [],

        minYear: currentYearVal - 5 ,
        maxYear: currentYearVal + 1 ,
        yearSlider: currentYearVal,

        rankSlider: 5,
        userSearch: '',

        fetchSlider: 0,
        fetchTotal: 0,
        fetchCount: 0,
        fetchPerComp: 100,
        fetchLabel: '',
        showProgress: false,
        batchData: createBatchData( null ),
  
  };
}


public componentDidMount() {

  this.updateWebInfo( this.state.parentWeb );
}

public async updateWebInfo ( webUrl?: string ) {

  console.log('_onWebUrlChange Fetchitng Lists ====>>>>> :', webUrl );

  let errMessage = null;
  let stateError : any[] = [];

  let pickedWeb = await getWebInfoIncludingUnique( webUrl, 'min', false, ' > GenWP.tsx ~ 825', 'BaseErrorTrace' );

  errMessage = pickedWeb.error;
  if ( pickedWeb.error && pickedWeb.error.length > 0 ) {
    stateError.push( <div style={{ padding: '15px', background: 'yellow' }}> <span style={{ fontSize: 'larger', fontWeight: 600 }}>Can't find the site</span> </div>);
    stateError.push( <div style={{ paddingLeft: '25px', paddingBottom: '30px', background: 'yellow' }}> <span style={{ fontSize: 'large', color: 'red'}}> { errMessage }</span> </div>);
  }

  let theSite: ISite = await getSiteInfo( webUrl, false, ' > GenWP.tsx ~ 831', 'BaseErrorTrace' );

  let listSelect = ['Title','ItemCount','LastItemUserModifiedDate','Created','BaseType','Id','DocumentTemplateUrl'].join(',');
  // let listSelect = ['*'].join(',');

  let thisWebInstance = Web(webUrl);

  const listObject = thisWebInstance.lists.getByTitle(this.state.listTitle);
  
  //https://github.com/pnp/pnpjs/issues/160#issuecomment-793849161
  listObject.query.set('t', new Date().getTime().toString()); // <-- forces unique Get path

  let theList: IEXStorageList = await listObject.select( listSelect ).get();
  // let theList: IEXStorageList = await thisWebInstance.lists.getByTitle(this.state.listTitle).get();
  theList.LibraryUrl = theList.DocumentTemplateUrl.replace('/Forms/template.dotx','/');

  let isCurrentWeb: boolean = false;
  if ( webUrl.toLowerCase().indexOf( this.props.pageContext.web.serverRelativeUrl.toLowerCase() ) > -1 ) { isCurrentWeb = true ; }

  let minYear: any = new Date( theList.Created);
  minYear = minYear.getFullYear();
  let maxYear: any = new Date( theList.LastItemUserModifiedDate);
  maxYear = maxYear.getFullYear() + 1;

  let currentYear: any = new Date();
  currentYear = currentYear.getFullYear();

  let currentUser = this.props.currentUser === null ? await this.getCurrentUser( this.props.parentWeb ) : this.props.currentUser;


  //Automatically kick off if it's under 5k items
  if ( theList.ItemCount > 0 && theList.ItemCount < 5000 ) {
    this.setState({ parentWeb: webUrl, stateError: stateError, pickedWeb: pickedWeb, isCurrentWeb: isCurrentWeb, theSite: theSite, currentUser: currentUser,
        pickedList: theList, fetchSlider: theList.ItemCount, minYear: minYear, maxYear: maxYear, yearSlider: currentYear });
    this.fetchStoredItems(pickedWeb, theList, theList.ItemCount, currentUser );

  } else {

    this.setState({ parentWeb: webUrl, stateError: stateError, pickedWeb: pickedWeb, isCurrentWeb: isCurrentWeb, theSite: theSite, currentUser: currentUser,
      pickedList: theList, isLoaded: true, isLoading: false, minYear: minYear, maxYear: maxYear, yearSlider: currentYear });

  }

  return;

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

  public render(): React.ReactElement<IExStorageProps> {

    let batchData = this.state.batchData;
    const batches = this.state.batches;

    
    let timeComment = null;
    let etaMinutes = this.state.pickedList && this.state.fetchSlider > 0 ? (  this.state.fetchSlider * 7 / ( 1000 * 60 ) ).toFixed( 1 ) : 0;
    if ( this.state.isLoading ) {
      timeComment = etaMinutes ;
    }
    let loadingNote = this.state.isLoading ? <div>
      Please do not interupt the process which could take { etaMinutes } minutes.
    </div> : null;
    let searchSpinner = this.state.isLoading ? 
        <Spinner size={SpinnerSize.large} label={` fetching ${this.state.pickedList ? this.state.fetchSlider : 'TBD'} items...`} />
     : null ;

    let myProgress = 1 === 1 ? <ProgressIndicator 
    label={ this.state.fetchLabel } 
    description={ '' } 
    percentComplete={ this.state.fetchPerComp } 
    progressHidden={ !this.state.showProgress }/> : null;


    let fetchButton = <PrimaryButton text={ 'Begin'} onClick={this.fetchStoredItemsClick.bind(this)} allowDisabledFocus disabled={ this.state.isLoading } />;

    let currentYear = new Date();
    let currentYearVal = currentYear.getFullYear();

    let sliderYearItself = !this.state.pickedList ? null : 
      <div style={{margin: '0 50px'}}> { createSlider( null , this.state.yearSlider , this.state.minYear, this.state.maxYear, 1 , this._updateMaxYear.bind(this), this.state.isLoading, 350) }</div> ;

    let sliderYearComponent = !this.state.pickedList ? null : <div className={ styles.inflexWrapCenter}>
      <span style={{ fontSize: 'larger', fontWeight: 'bolder', minWidth: '300px' }}> { `Ignore files Created > ${ this.state.yearSlider }` } </span>
      { sliderYearItself }
    </div>;

    let sliderCountItself = !this.state.pickedList ? null : 
      <div style={{margin: '0 50px'}}> { createSlider( null , this.state.fetchSlider , 0, this.state.pickedList.ItemCount, batchSize , this._updateMaxFetch.bind(this), this.state.isLoading, 350) }</div> ;

    let sliderCountComponent = !this.state.pickedList ? null : <div className={ styles.inflexWrapCenter}>
      <span style={{ fontSize: 'larger', fontWeight: 'bolder', minWidth: '300px' }}> { `Fetch up to ${this.state.pickedList.ItemCount } Files` } </span>
      { sliderCountItself }
      <span style={{marginRight: '50px'}}> { `Plan for about ${etaMinutes} minutes` } </span>
      { fetchButton }
    </div>;

    let typesPivotContent = <div>
      <ExTypes 
  
          //Size courtesy of https://www.netwoven.com/2018/11/13/resizing-of-spfx-react-web-parts-in-different-scenarios/
          WebpartHeight = { this.props.WebpartHeight }
          WebpartWidth = { this.props.WebpartWidth }
      
          pickedWeb  = { this.state.pickedWeb }
          pickedList = { this.state.pickedList }
          theSite = {null }

          typesInfo = { batchData.typesInfo }
          batches = { batches }
          batchData = { batchData }
      >
      </ExTypes>
      <ReactJson src={ batchData.typesInfo.types } name={ 'Types' } collapsed={ true } displayDataTypes={ true } displayObjectSize={ true } enableClipboard={ true } style={{ padding: '20px 0px' }}/></div>;

    let usersPivotContent = null;
    
    if ( batchData.userInfo !== null ) {
      let rankSlider = this.state.rankSlider;
      let sliderMin = batchData.userInfo.allUsers.length < 3 ? batchData.userInfo.allUsers.length : 3;
      let sliderUserCount = batchData.userInfo.allUsers.length < 5 ? null : 
        <div style={{margin: '0px 50px 20px 50px'}}> { createSlider( 'Show Top' , rankSlider , sliderMin, batchData.userInfo.allUsers.length, 1 , this._updateRankShow.bind(this), this.state.isLoading, 350) }</div> ;
      
      usersPivotContent = <div><div>
        <h3>Summary of files by user</h3>
        <div className={ styles.inflexWrapCenter}>
          <div> { sliderUserCount } </div>
          <div> { this.buildSearchBox() } </div>
        </div>

        <div className={ styles.inflexWrapCenter}>
          { this.buildUserTables( batchData.userInfo.createSizeRank, batchData.userInfo.allUsers, 'createSizeRank', rankSlider, this.state.userSearch ) }
          { this.buildUserTables( batchData.userInfo.createCountRank, batchData.userInfo.allUsers, 'createCountRank', rankSlider, this.state.userSearch ) }
          { this.buildUserTables( batchData.userInfo.modifySizeRank, batchData.userInfo.allUsers, 'modifySizeRank', rankSlider, this.state.userSearch ) }
          { this.buildUserTables( batchData.userInfo.modifyCountRank, batchData.userInfo.allUsers, 'modifyCountRank', rankSlider, this.state.userSearch )}
        </div>
        {/* <p> { batchData.userInfo.allUsers.map( user => { return user.userTitle; }).join(', ') }</p> */}
      </div>
      <ReactJson src={ batchData.userInfo.allUsers } name={ 'Users' } collapsed={ true } displayDataTypes={ true } displayObjectSize={ true } enableClipboard={ true } style={{ padding: '20px 0px' }}/></div>;
    }

    let sizePivotContent = <div><div>
      <h3>Summary of files by Size</h3>
      </div>
        <ReactJson src={ batchData.large.summary } name={ `Summary` } 
            collapsed={ true } displayDataTypes={ true } displayObjectSize={ true } enableClipboard={ true } style={{ padding: '20px 0px' }}/>
        <ReactJson src={ batchData.large.GT10G } name={ '> 10GB per file' } collapsed={ true } displayDataTypes={ true } displayObjectSize={ true } enableClipboard={ true } style={{ padding: '20px 0px' }}/>
        <ReactJson src={ batchData.large.GT01G } name={ '> 1GB per file' } collapsed={ true } displayDataTypes={ true } displayObjectSize={ true } enableClipboard={ true } style={{ padding: '20px 0px' }}/>
        <ReactJson src={ batchData.large.GT100M } name={ '> 100MB per file' } collapsed={ true } displayDataTypes={ true } displayObjectSize={ true } enableClipboard={ true } style={{ padding: '20px 0px' }}/>
        <ReactJson src={ batchData.large.GT10M } name={ '> 10M per file' } collapsed={ true } displayDataTypes={ true } displayObjectSize={ true } enableClipboard={ true } style={{ padding: '20px 0px' }}/>
      </div>;

    let agePivotContent = <div><div>
      <h3>Summary of files by Age</h3>
      </div>
        <ReactJson src={ batchData.oldCreated.summary } name={ `Summary` } 
          collapsed={ true } displayDataTypes={ true } displayObjectSize={ true } enableClipboard={ true } style={{ padding: '20px 0px' }}/>

        <ReactJson src={ batchData.oldCreated.Age5Yr } name={ `Created before ${ (this.currentYear -4 ) }` } 
          collapsed={ true } displayDataTypes={ true } displayObjectSize={ true } enableClipboard={ true } style={{ padding: '20px 0px' }}/>

        <ReactJson src={ batchData.oldCreated.Age4Yr } name={ `Created in ${ (this.currentYear -4 ) }` } 
          collapsed={ true } displayDataTypes={ true } displayObjectSize={ true } enableClipboard={ true } style={{ padding: '20px 0px' }}/>

        <ReactJson src={ batchData.oldCreated.Age3Yr } name={ `Created in ${ (this.currentYear -3 ) }` } 
          collapsed={ true } displayDataTypes={ true } displayObjectSize={ true } enableClipboard={ true } style={{ padding: '20px 0px' }}/>

        <ReactJson src={ batchData.oldCreated.Age2Yr } name={ `Created in ${ (this.currentYear -2 ) }` } 
          collapsed={ true } displayDataTypes={ true } displayObjectSize={ true } enableClipboard={ true } style={{ padding: '20px 0px' }}/>

        <ReactJson src={ batchData.oldCreated.Age1Yr } name={ `Created in ${ (this.currentYear -1 ) }` } 
          collapsed={ true } displayDataTypes={ true } displayObjectSize={ true } enableClipboard={ true } style={{ padding: '20px 0px' }}/>

      </div>;


    let youPivotContent = <div>
        <ExUser 
          pageContext = { this.context.pageContext }
          wpContext = { this.context }
          tenant = { this.props.tenant }
      
          //Size courtesy of https://www.netwoven.com/2018/11/13/resizing-of-spfx-react-web-parts-in-different-scenarios/
          WebpartElement = { this.props.WebpartElement }
  
          //Size courtesy of https://www.netwoven.com/2018/11/13/resizing-of-spfx-react-web-parts-in-different-scenarios/
          WebpartHeight = { this.props.WebpartHeight }
          WebpartWidth = { this.props.WebpartWidth }
      
          pickedWeb  = { this.state.pickedWeb }
          pickedList = { this.state.pickedList }
          theSite = {null }
  
          isLoaded = {false }
      
          currentUser = {this.state.currentUser }
          isCurrentUser = { true }
          userSummary = { batchData.userInfo.currentUser }
          batches = { batches }
          batchData = { batchData }
        >
        </ExUser>
      </div>;

    let permsPivotContent = <div><div>
      <h3>Summary of files with broken permissions</h3>
      </div>
        <ReactJson src={ batchData.uniqueInfo.uniqueRolls} name={ 'Broken Permissions' } collapsed={ true } displayDataTypes={ true } displayObjectSize={ true } enableClipboard={ true } style={{ padding: '20px 0px' }}/>
      </div>;

    let dupsPivotContent = <div><div>
      <h3>Summary of duplicate files</h3>
      </div>
        <ReactJson src={ batchData.duplicateInfo.duplicates} name={ 'Duplicate files' } collapsed={ true } displayDataTypes={ true } displayObjectSize={ true } enableClipboard={ true } style={{ padding: '20px 0px' }}/>
      </div>;

    let folderPivotContent = <div><div>
      <h3>Summary of Folders</h3>
      </div>
        <ReactJson src={ batchData.folderInfo.folders} name={ 'Folders' } collapsed={ true } displayDataTypes={ true } displayObjectSize={ true } enableClipboard={ true } style={{ padding: '20px 0px' }}/>
      </div>;

    let componentPivot = 
    <Pivot
        styles={ pivotStyles }
        linkFormat={PivotLinkFormat.tabs}
        linkSize={PivotLinkSize.normal}
        // onLinkClick={this._selectedListDefIndex.bind(this)}
    > 
      <PivotItem headerText={ pivotHeading1 } ariaLabel={pivotHeading1} title={pivotHeading1} itemKey={ pivotHeading1 } keytipProps={ { content: 'Hello', keySequences: ['a','b','c'] } }>
        <ReactJson src={ batchData } name={ 'Summary' } collapsed={ true } displayDataTypes={ true } displayObjectSize={ true } enableClipboard={ true } style={{ padding: '20px 0px' }}/>
      </PivotItem>

      <PivotItem headerText={ pivotHeading2 } ariaLabel={pivotHeading2} title={pivotHeading2} itemKey={ pivotHeading2 } keytipProps={ { content: 'Hello', keySequences: ['a','b','c'] } }>
        { typesPivotContent }
      </PivotItem>

      <PivotItem headerText={ pivotHeading3 } ariaLabel={pivotHeading3} title={pivotHeading3} itemKey={ pivotHeading3 } keytipProps={ { content: 'Hello', keySequences: ['a','b','c'] } }>
        { usersPivotContent }
      </PivotItem>

      <PivotItem headerText={ pivotHeading4 } ariaLabel={pivotHeading4} title={pivotHeading4} itemKey={ pivotHeading4 } keytipProps={ { content: 'Hello', keySequences: ['a','b','c'] } }>
        { sizePivotContent }
      </PivotItem> 
      
      <PivotItem headerText={ pivotHeading5 } ariaLabel={pivotHeading5} title={pivotHeading5} itemKey={ pivotHeading5 } keytipProps={ { content: 'Hello', keySequences: ['a','b','c'] } }>
        { agePivotContent }
      </PivotItem>

      <PivotItem headerText={ pivotHeading6 } ariaLabel={pivotHeading6} title={pivotHeading6} itemKey={ pivotHeading6 } keytipProps={ { content: 'Hello', keySequences: ['a','b','c'] } }>
        { youPivotContent }
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
    </Pivot>;

    let userPanel = null;
    
    if ( this.state.showUser > -1 ) { 
      let panelContent = <ExUser 
        pageContext = { this.context.pageContext }
        wpContext = { this.context }
        tenant = { this.props.tenant }
    
        //Size courtesy of https://www.netwoven.com/2018/11/13/resizing-of-spfx-react-web-parts-in-different-scenarios/
        WebpartElement = { this.props.WebpartElement }

        //Size courtesy of https://www.netwoven.com/2018/11/13/resizing-of-spfx-react-web-parts-in-different-scenarios/
        WebpartHeight = { this.props.WebpartHeight }
        WebpartWidth = { this.props.WebpartWidth }
    
        pickedWeb  = { this.state.pickedWeb }
        pickedList = { this.state.pickedList }
        theSite = {null }

        isLoaded = {false }
    
        currentUser = {this.state.currentUser }
        isCurrentUser = { true }
        userSummary = { batchData.userInfo.allUsers[ this.state.showUser ] }
        batches = { batches }
        batchData = { batchData }
      >
    </ExUser>;

    userPanel = <div><Panel
      isOpen={ this.state.showUser > -1 ? true : false }
      // this prop makes the panel non-modal
      isBlocking={true}
      onDismiss={ this._onCloseUser.bind(this) }
      closeButtonAriaLabel="Close"
      type = { PanelType.large }
      isLightDismiss = { true }
      >
        { panelContent }
    </Panel></div>;

    }
    return (
      <div className={ styles.exStorage }>
        <div className={ styles.container }>

          {/* <span className={ styles.title }>Welcome to SharePoint!</span> */}
          {/* <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p> */}
          <p className={ styles.description }>{escape(this.props.parentWeb)}</p>
          {/* <div>{ this.state.currentUser ? this.state.currentUser.Title : null }</div> */}

          { sliderYearComponent }
          { sliderCountComponent }
          { this.state.isLoading ? 
              <div>
                { loadingNote }
                { searchSpinner }
                { myProgress }
              </div>
            : null
          } 

          { componentPivot }
          <div style={{ overflowY: 'auto' }}>
              {/* <ReactJson src={ this.state.currentUser } collapsed={ true } displayDataTypes={ true } displayObjectSize={ true } enableClipboard={ true } style={{ padding: '20px 0px' }}/>
              <ReactJson src={ this.state.pickedWeb } collapsed={ true } displayDataTypes={ true } displayObjectSize={ true } enableClipboard={ true } style={{ padding: '20px 0px' }}/>
              <ReactJson src={ this.state.pickedList } collapsed={ true } displayDataTypes={ true } displayObjectSize={ true } enableClipboard={ true } style={{ padding: '20px 0px' }}/> */}
              <ReactJson src={ batches } name={ 'All items' } collapsed={ true } displayDataTypes={ true } displayObjectSize={ true } enableClipboard={ true } style={{ padding: '20px 0px' }}/>

          </div>
          { userPanel }
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
        { `Search all ${ this.state.batchData.userInfo.allUsers.length } users` }
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

    this.setState({ userSearch: item });
  }

  private buildUserTables( indexs: number[], users: IUserSummary[] , data: string, countToShow: number, userSearch: string ): any {

    // { this.buildUserTables( batchData.userRanks.createSizeRank, batchData.userInfo.allUsers, 'createSizeRank') }
    // { this.buildUserTables( batchData.userRanks.createCountRank, batchData.userInfo.allUsers, 'createCountRank') }
    // { this.buildUserTables( batchData.userRanks.modifySizeRank, batchData.userInfo.allUsers, 'modifySizeRank') }
    // { this.buildUserTables( batchData.userRanks.modifyCountRank, batchData.userInfo.allUsers, 'modifyCountRank') }

    let elements = [];
    let tableTitle = data;

    indexs.map( ( allUserIndex, index ) => {

      if ( index < countToShow || userSearch.length > 0 ) {
        let user = users[ allUserIndex ];
        let label = '' ;
        let createTotalSizeLabel = user.createTotalSizeLabel;
        let modifyTotalSizeLabel = user.modifyTotalSizeLabel;
        let createPercent = ( user.summary.sizeP ).toFixed( 0 );
        let countPercent = ( user.summary.countP ).toFixed( 0 );

        switch (data) {
          case 'createSizeRank':
            label = `${user.userTitle}  [ ${ createTotalSizeLabel } / ${ createPercent }% ]` ;
            break;
          case 'createCountRank':
            label = `${user.userTitle}  [ ${user.createCount} / ${ countPercent }% ]` ;
            break;
          case 'modifySizeRank':
            label = `${user.userTitle}  [ ${ modifyTotalSizeLabel } ]` ;
            break;
          case 'modifyCountRank':
            label = `${user.userTitle}  [ ${user.modifyCount} ]` ;
            break;
  
          default:
            break;
        }
        
        let title = `( #${ allUserIndex } Id: ${user.userId} ) ${user.userTitle} created ${ createTotalSizeLabel }, modified ${modifyTotalSizeLabel}` ;

        let showListItem = userSearch.length === 0 || (userSearch.length > 0 && user.userTitle.toLowerCase().indexOf(userSearch.toLowerCase() )  > -1  ) ? true : false;

        let liStyle : React.CSSProperties = showListItem === true ?
        {
          display: 'flex',
          flexDirection: 'row',
          justifyContent: 'flex-start',
          alignItems: 'center',
        } : { display: 'none' };

        elements.push(<li title={ title} style= { liStyle } onClick={ this._onClickUser.bind(this)} id={ allUserIndex.toFixed(0) }>
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
    
  private _onClickUser( event ){
    console.log( event );
    console.log( event.currentTarget.id );
    let showUserId = parseInt(event.currentTarget.id);
    this.setState({
      showUser: showUserId,
    });
  }
    
  private _onCloseUser( event ){
    this.setState({
      showUser: -1,
    });
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

  private fetchStoredItemsClick( ) {
    this.fetchStoredItems( this.state.pickedWeb, this.state.pickedList, this.state.fetchSlider, this.state.currentUser );
  }

  private fetchStoredItems( pickedWeb: IPickedWebBasic , pickedList: IEXStorageList, getCount: number, currentUser: IUser ) {

    this.setState({ 
      isLoading: true,
      errorMessage: '',
    });
    getSearchedFiles( this.props.tenant, pickedList, true);
    getStorageItems( pickedWeb, pickedList, getCount, currentUser, this.addTheseItemsToState.bind(this), this.setProgress.bind(this) );

  }

  private addTheseItemsToState ( batchInfo ) {

    // let isLoading = this.props.showPrefetchedPermissions === true ? false : myPermissions.isLoading;
    // let showNeedToWait = this.state.showNeedToWait === false ? false :
    //   isLoading === true ?  true : false;


    console.log('addTheseItemsToState');
    this.setState({ 
      // items: batch.items,
      
      isLoading: false,
      // showNeedToWait: false,

      // errorMessage: batch.errMessage,
      // hasError: batch.errMessage.length > 0 ? true : false,

      // fetchTotal: fetchTotal,
      // fetchCount: fetchCount,
      // fetchPerComp: fetchPerComp,
      // fetchLabel: fetchLabel,
      showProgress: false,

      batches: batchInfo.batches,
      batchData: batchInfo.batchData,


    });

  }

  private setProgress( fetchCount, fetchTotal, fetchLabel ) {
    let fetchPerComp = fetchTotal > 0 ? fetchCount / fetchTotal : 0 ;
    let showProgress = fetchCount !== fetchTotal ? true : false;

    this.setState({
      fetchTotal: fetchTotal,
      fetchCount: fetchCount,
      fetchPerComp: fetchPerComp,
      fetchLabel: fetchLabel,
      showProgress: showProgress,
    });

  }


  public async getCurrentUser( webURL: string): Promise<IUser> {
    let currentUser : IUser =  null;
    let thisWebInstance = Web(webURL);
    await thisWebInstance.currentUser.get().then((r) => {
      currentUser = {
        title: r['Title'] , //
        Title: r['Title'] , //
        initials: r['Title'].split(" ").map((n)=>n[0]).join(""), //Single person column
        email: r['Email'] , //Single person column
        id: r['Id'] , //
        Id: r['Id'] , //
        ID: r['Id'] , //
        remoteID: null,
        isSiteAdmin: r['IsSiteAdmin'],
        LoginName: r['LoginName'],
        Name: r['LoginName'],
      };
      // this.setState({ currentUser: currentUser });
      
    });
    return currentUser;
  }

}
