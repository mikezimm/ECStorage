import * as React from 'react';
import styles from '../../ExStorage.module.scss';
import itemStyles from './Items.module.scss';

import { IEsItemsProps } from './IEsItemsProps';
import { IEsItemsState } from './IEsItemsState';
import { IExStorageState, IEXStorageList, IEXStorageBatch, IBatchData, IUserSummary, IFileType, IItemDetail, IDuplicateFile, IItemType, IFolderDetail, IKnownMeta, ISharingInfo,  } from '../../IExStorageState';
import { escape } from '@microsoft/sp-lodash-subset';


import { sp, Views, IViews, ISite } from "@pnp/sp/presets/all";
import { Web, IList, Site } from "@pnp/sp/presets/all";

import { ProgressIndicator } from 'office-ui-fabric-react/lib/ProgressIndicator';

import {
  Spinner,
  SpinnerSize,
  FloatingPeoplePicker,
  tdProperties,
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
import { SearchBox, ISearchBoxProps } from 'office-ui-fabric-react/lib/SearchBox';
import { Icon  } from 'office-ui-fabric-react/lib/Icon';

import { Panel, IPanelProps, IPanelStyleProps, IPanelStyles, PanelType } from 'office-ui-fabric-react/lib/Panel';

import { Pivot, PivotItem, IPivotItemProps, PivotLinkFormat, PivotLinkSize,} from 'office-ui-fabric-react/lib/Pivot';
import { Dropdown, DropdownMenuItemType, IDropdownStyles, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { MessageBar, MessageBarType,  } from 'office-ui-fabric-react/lib/MessageBar';

import { IFrameDialog,  } from "@pnp/spfx-controls-react/lib/IFrameDialog";
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';

import ReactJson from "react-json-view";

import { IPickedWebBasic, IPickedList, }  from '@mikezimm/npmfunctions/dist/Lists/IListInterfaces';
import { IUser } from '@mikezimm/npmfunctions/dist/Services/Users/IUserInterfaces';

import { getSiteInfo, getWebInfoIncludingUnique } from '@mikezimm/npmfunctions/dist/Services/Sites/getSiteInfo';
import { cleanURL } from '@mikezimm/npmfunctions/dist/Services/Strings/urlServices';
import { buildAppWarnIcon, buildClickableIcon } from '@mikezimm/npmfunctions/dist/Icons/stdIconsBuildersV02';

import * as StdIcons from '@mikezimm/npmfunctions/dist/Icons/iconNames';

// import { sortObjectArrayByChildNumberKey, sortNumberArray } from '@mikezimm/npmfunctions/dist/Services/Arrays/sorting';
import { sortObjectArrayByChildNumberKey, } from '@mikezimm/npmfunctions/dist/Services/Arrays/sorting';

import * as fpsAppIcons from '@mikezimm/npmfunctions/dist/Icons/standardExStorage';

import { createSlider, createChoiceSlider } from '../../fields/sliderFieldBuilder';

import { getSizeLabel, getCountLabel, getCommaSepLabel } from '@mikezimm/npmfunctions/dist/Services/Math/basicOperations';

import { getSearchedFiles } from '../../ExSearch';

import { createItemsHeadingWithTypeIcons } from '../miniComps/components';

import { createItemDetail, getItemSearchString, getEventSearchString, getHighlightedText } from './SingleItem';

import { IItemSharingInfo, ISharingEvent, ISharedWithUser } from '../../Sharing/ISharingInterface';


/***
 *     .o88b.  .d88b.  d8b   db .d8888. d888888b 
 *    d8P  Y8 .8P  Y8. 888o  88 88'  YP `~~88~~' 
 *    8P      88    88 88V8o 88 `8bo.      88    
 *    8b      88    88 88 V8o88   `Y8b.    88    
 *    Y8b  d8 `8b  d8' 88  V888 db   8D    88    
 *     `Y88P'  `Y88P'  VP   V8P `8888Y'    YP    
 *                                               
 *                                               
 */

export default class EsItems extends React.Component<IEsItemsProps, IEsItemsState> {

  private showHeading: boolean = this.props.showHeading === false ? false : true;
  private currentDate = new Date();
  private currentYear = this.currentDate.getFullYear();

  private itemsAny: any[] = this.props.itemType === 'Items' ? this.props.items : this.props.itemType === 'Shared' ? this.props.sharedItems: this.props.duplicateInfo.duplicates;

  private sharedEvents: ISharingEvent[] = this.props.itemType === 'Shared' ? this.buildAllSharedEventsItems( this.props.sharedItems ): [];

  private itemsLength = this.props.itemType === 'Shared' ? this.sharedEvents.length : this.itemsAny.length;

  private getRelativePath = this.props.itemType === 'Items' && this.itemsLength > 0 ? true : false;
  private commonFolders: string[] = this.getRelativePath === true ? this.getCommonFolders( this.itemsAny ) : [];
  private commonRelativePath: string = this.getRelativePath === true && this.commonFolders.length > 0 ? this.commonFolders.join('/') : '';

  private checkedOut : boolean = this.props.heading.indexOf( 'Checked Out' ) > -1 ? true : false;

  private commonPath: string = this.getRelativePath === true && this.commonFolders.length > 0 ? this.commonRelativePath.replace( this.props.pickedList.LibraryUrl , '') + '/' : '';
  private commonParent: string = this.getRelativePath === true && this.commonFolders.length > 0 && this.commonRelativePath !== this.props.pickedList.LibraryUrl ? this.commonFolders[ this.commonFolders.length - 1 ] : '';

  private itemsHeading: any = createItemsHeadingWithTypeIcons( this.itemsAny, this.props.itemType, this.props.heading, this.props.icons, this._onClickType.bind(this) );

  private sliderTitle = this.itemsLength < 400 ? 'Show Top items by size' : `Show up to 400 of ${ getCommaSepLabel(this.itemsLength) } items, use Search box to find more)`;
  private sliderMax = this.itemsLength < 400 ? this.itemsLength : 400;
  private sliderInc = this.itemsLength < 50 ? 1 : this.itemsLength < 100 ? 10 : 25;
  private siderMin = this.sliderInc > 1 ? this.sliderInc : 5;

  private searchMedia = this.props.dataOptions.useMediaTags !== true ? '' : ', MediaServiceAutoTags, MediaServiceKeyPoints, MediaServiceLocation, MediaServiceOCR';
  private searchNote = `Search will search Created Name and Date, filenames/types ${ this.searchMedia }`;
  private visibleNote = `You will find an item under a User if the User created the item.`;

  private getCommonFolders( itemsIn: IItemDetail[] | IDuplicateFile[] ) {
    let items: any[] = itemsIn;

    if ( itemsIn.length === 0 ) { return []; }

    let commonFolders: string[] = items[0].parentFolder.split('/');
    let startTime = new Date();

    items.map( item => {
      let itemFolders: string[] = item.parentFolder.split('/');
      let newCommonFolders : string[] = [];
      commonFolders.map( ( folder, index ) => {
        //If current folder of item is the same as the path of the commonFolders, then push it
        if ( folder === itemFolders [ index ]  ) { newCommonFolders.push( folder ) ; } 
      });
      commonFolders = newCommonFolders;
    } );
    let endTime = new Date();
    let processTime = endTime.getTime() - startTime.getTime();
    console.log('processTime(s), commonFolders: ', processTime / 1000, commonFolders );

    return commonFolders;
  }

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



public constructor(props:IEsItemsProps){
  super(props);

  let currentYear = new Date();
  let currentYearVal = currentYear.getFullYear();
  
  let totalSize: number = 0;
  this.itemsAny.map( item => {
    totalSize += item.size;
  });

  
  if ( this.props.itemType === 'Duplicates' ) {
    this.searchNote = `Search will search Created Name and Date, foldername/types`;
    this.visibleNote = `Folder names start at the lowest common branch (folder)`;

  } else if ( this.props.itemType === 'Shared' ) {
    this.searchNote = `Search will search date/time, sharedBy, sharedWith, filenames`;
    this.visibleNote = `Shared items under a user include any files they created, modified, shared or were shared with.`;

  }

  this.state = {

        isLoaded: true,
        isLoading: false,
        errorMessage: '',

        hasError: false,
      
        showPane: false,

        items: [],
        totalSize: totalSize,
        
        showItems: [],
        dups: this.props.itemType === 'Duplicates' ? this.props.duplicateInfo.duplicates : [],

        minYear: currentYearVal - 5 ,
        maxYear: currentYearVal + 1 ,

        rankSlider: this.siderMin,
        textSearch: '',

        fetchSlider: 0,
        fetchTotal: 0,
        fetchCount: 0,
        showProgress: false,
        fetchPerComp: 100,
        fetchLabel: '',

        filteredCount: this.itemsLength,

        showItem: false,
        showPreview: false,
        selectedItem: null,

        hasMedia: false,
  
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


  /***
 *    d8888b. d88888b d8b   db d8888b. d88888b d8888b. 
 *    88  `8D 88'     888o  88 88  `8D 88'     88  `8D 
 *    88oobY' 88ooooo 88V8o 88 88   88 88ooooo 88oobY' 
 *    88`8b   88~~~~~ 88 V8o88 88   88 88~~~~~ 88`8b   
 *    88 `88. 88.     88  V888 88  .8D 88.     88 `88. 
 *    88   YD Y88888P VP   V8P Y8888D' Y88888P 88   YD 
 *                                                     
 *                                                     
 */
  public render(): React.ReactElement<IEsItemsProps> {

    console.log('EsItems.tsx1');
    // debugger;
    // const items : IItemDetail[] | IDuplicateFile []= this.itemsAny;
    let itemsTable = null;
    let filteredCount = null;

    if ( this.props.itemType === 'Shared' ) {
      let tableResults = this.buildSharedEventsTable( this.sharedEvents , '', this.state.rankSlider, this.state.textSearch );
      itemsTable = tableResults.table;
      filteredCount = tableResults.filteredCount;

    } else if ( this.props.itemType === 'Items' ) {
      let tableResults = this.buildItemsTable( this.props.items , this.props.itemsAreDups, '', this.state.rankSlider, this.state.textSearch, 'size' );
      itemsTable = tableResults.table;
      filteredCount = tableResults.filteredCount;

    } else if ( this.props.itemType === 'Duplicates' ) {
      let tableResults = this.buildDupsTable( this.props.duplicateInfo.duplicates , this.props.itemsAreDups, '', this.state.rankSlider, this.state.textSearch, 'size' );
      itemsTable = tableResults.table;
      filteredCount = tableResults.filteredCount;

    }

    let page = null;
    let userPanel = null;

    const emptyItemsElements = this.props.emptyItemsElements;

    let showEmptyElements = false;
    if ( this.props.itemType === 'Items' ) {
      showEmptyElements = this.props.items.length === 0 && emptyItemsElements && emptyItemsElements.length > 0 ? true : false;

    } else if ( this.props.itemType === 'Duplicates') {
      showEmptyElements = this.props.duplicateInfo.duplicates.length === 0 && emptyItemsElements && emptyItemsElements.length > 0 ? true : false;

    } else if ( this.props.itemType === 'Shared') {
      showEmptyElements = this.props.sharedItems.length === 0 && emptyItemsElements && emptyItemsElements.length > 0 ? true : false;
      
    }

    let component = <div className={ styles.inflexWrapCenter}>
      { itemsTable }
    </div>;

    let sliderTypeCount = this.itemsLength < 5 ? null : 
      <div style={{margin: '0px 50px 20px 50px'}}> { createSlider( this.sliderTitle , this.state.rankSlider , this.siderMin, this.sliderMax, this.sliderInc , this._typeSlider.bind(this), this.state.isLoading, 350) }</div> ;

    if ( showEmptyElements ) {
      page = emptyItemsElements[Math.floor(Math.random()*emptyItemsElements.length)];  //https://stackoverflow.com/a/5915122

    } else {

      let panelContent = null;

/***
 *    d888888b        d88888b d8888b.  .d8b.  .88b  d88. d88888b      d8888b. d888888b  .d8b.  db       .d88b.   d888b  
 *      `88'          88'     88  `8D d8' `8b 88'YbdP`88 88'          88  `8D   `88'   d8' `8b 88      .8P  Y8. 88' Y8b 
 *       88           88ooo   88oobY' 88ooo88 88  88  88 88ooooo      88   88    88    88ooo88 88      88    88 88      
 *       88    C8888D 88~~~   88`8b   88~~~88 88  88  88 88~~~~~      88   88    88    88~~~88 88      88    88 88  ooo 
 *      .88.          88      88 `88. 88   88 88  88  88 88.          88  .8D   .88.   88   88 88booo. `8b  d8' 88. ~8~ 
 *    Y888888P        YP      88   YD YP   YP YP  YP  YP Y88888P      Y8888D' Y888888P YP   YP Y88888P  `Y88P'   Y888P  
 *                                                                                                                      
 *                                                                                                                      
 */

      if ( this.state.showPreview === true && this.state.selectedItem ) {



        userPanel = <IFrameDialog 
          url={this.state.selectedItem.ServerRedirectedEmbedUrl}
          // iframeOnLoad={this._onIframeLoaded.bind(this)}
          hidden={ false }
          onDismiss={this._onDialogDismiss.bind(this)}
          modalProps={{
              isBlocking: true,
              // containerClassName: styles.dialogContainer
          }}
          dialogContentProps={{
              type: DialogType.close,
              showCloseButton: true
          }}
          onDismissed= { this._onDialogDismiss.bind( this ) }
          width={'60%'}
          height={'60%'}/>;

/***
 *    d8888b. d88888b d888888b  .d8b.  d888888b db           d8888b.  .d8b.  d8b   db d88888b db      
 *    88  `8D 88'     `~~88~~' d8' `8b   `88'   88           88  `8D d8' `8b 888o  88 88'     88      
 *    88   88 88ooooo    88    88ooo88    88    88           88oodD' 88ooo88 88V8o 88 88ooooo 88      
 *    88   88 88~~~~~    88    88~~~88    88    88           88~~~   88~~~88 88 V8o88 88~~~~~ 88      
 *    88  .8D 88.        88    88   88   .88.   88booo.      88      88   88 88  V888 88.     88booo. 
 *    Y8888D' Y88888P    YP    YP   YP Y888888P Y88888P      88      YP   YP VP   V8P Y88888P Y88888P 
 *                                                                                                    
 *                                                                                                    
 */

      } else if ( this.state.selectedItem ) { 


        panelContent = createItemDetail( this.state.selectedItem, this.props.itemsAreDups, this.props.pickedWeb.url, this.state.textSearch, this._onCloseItemDetail.bind( this ), this._onPreviewClick.bind( this ) );
    
/***
 *    .d8888. db   db  .d88b.  db   d8b   db      d888888b d888888b d88888b .88b  d88. .d8888. 
 *    88'  YP 88   88 .8P  Y8. 88   I8I   88        `88'   `~~88~~' 88'     88'YbdP`88 88'  YP 
 *    `8bo.   88ooo88 88    88 88   I8I   88         88       88    88ooooo 88  88  88 `8bo.   
 *      `Y8b. 88~~~88 88    88 Y8   I8I   88         88       88    88~~~~~ 88  88  88   `Y8b. 
 *    db   8D 88   88 `8b  d8' `8b d8'8b d8'        .88.      88    88.     88  88  88 db   8D 
 *    `8888Y' YP   YP  `Y88P'   `8b8' `8d8'       Y888888P    YP    Y88888P YP  YP  YP `8888Y' 
 *                                                                                             
 *                                                                                             
 */

      } else if ( this.state.showItems.length > 0 ) {


        panelContent = <div style={{ marginTop: '1em' }}>
          <EsItems 
            pickedWeb  = { this.props.pickedWeb }
            pickedList = { this.props.pickedList }
            theSite = {null }

            items = { this.state.showItems }
            itemsAreDups = { this.props.childrenAreDups ? this.props.childrenAreDups : false }
            itemsAreFolders = { false }
            duplicateInfo = { null }
            heading = { ` Duplicates of ${ this.state.showItems[0].FileLeafRef  }` }
            // batches = { batches }
            icons = { [ ]}
            emptyItemsElements = { emptyItemsElements }
              
            dataOptions = { this.props.dataOptions }
            uiOptions = { this.props.uiOptions }

            sharedItems = { [] }

            itemType = { 'Items' }

          ></EsItems>
        </div>;
      
      }

      
/***
 *    d8888b.  .d8b.  d8b   db d88888b db      
 *    88  `8D d8' `8b 888o  88 88'     88      
 *    88oodD' 88ooo88 88V8o 88 88ooooo 88      
 *    88~~~   88~~~88 88 V8o88 88~~~~~ 88      
 *    88      88   88 88  V888 88.     88booo. 
 *    88      YP   YP VP   V8P Y88888P Y88888P 
 *                                             
 *                                             
 */

      if ( panelContent !== null ) {
          userPanel = <div><Panel
          isOpen={ this.state.showItem === true || this.state.showItems.length > 0 ? true : false }
          // this prop makes the panel non-modal
          isBlocking={true}
          onDismiss={ this._onCloseItemDetail.bind(this) }
          closeButtonAriaLabel="Close"
          type = { PanelType.large }
          isLightDismiss = { true }
          >
            { panelContent }
        </Panel></div>;
      }

/***
 *    d8888b.  .d8b.   d888b  d88888b 
 *    88  `8D d8' `8b 88' Y8b 88'     
 *    88oodD' 88ooo88 88      88ooooo 
 *    88~~~   88~~~88 88  ooo 88~~~~~ 
 *    88      88   88 88. ~8~ 88.     
 *    88      YP   YP  Y888P  Y88888P 
 *                                    
 *                                    
 */

      let panelStyle = this.showHeading !== true ? { marginTop: '1.4em'} : null;

      let styleCommonPathDisplay = this.commonPath === '' ? 'none' : null ;

      let foundMessage = this.state.textSearch === '' ? 
        `All ${ this.itemsLength } items are below this folder: ${ this.props.pickedList.LibraryUrl.replace( this.props.pickedWeb.ServerRelativeUrl, '') }`: 
        `Found ${ filteredCount } items below this folder: ${ this.props.pickedList.LibraryUrl.replace( this.props.pickedWeb.ServerRelativeUrl, '') }`;

      page = <div style= { panelStyle }>
        { this.showHeading !== true ? null : this.itemsHeading }
        <div className={ styles.inflexWrapCenter}>
          <div> { sliderTypeCount } </div>
          <div> { this.buildSearchBox( this.state.textSearch) } </div>
        </div>
        <div>
          <div>{ this.searchNote }</div>
          <div>{ this.visibleNote }</div>
          <div style={{ padding: '10px 0px 5px 0px', display: styleCommonPathDisplay }}>
            { foundMessage }
            <span style={{ fontWeight: 600 }}>{ this.commonPath }</span></div>
        </div>
        { component }
      </div>;
    }



    /***
 *    d8888b. d88888b d888888b db    db d8888b. d8b   db 
 *    88  `8D 88'     `~~88~~' 88    88 88  `8D 888o  88 
 *    88oobY' 88ooooo    88    88    88 88oobY' 88V8o 88 
 *    88`8b   88~~~~~    88    88    88 88`8b   88 V8o88 
 *    88 `88. 88.        88    88b  d88 88 `88. 88  V888 
 *    88   YD Y88888P    YP    ~Y8888P' 88   YD VP   V8P 
 *                                                       
 *                                                       
 */
    return (
      <div className={ [styles.exStorage, itemStyles.itemsPage].join(' ') } style={{ marginLeft: '25px'}}>
        { page }
        { userPanel }
      </div>
    );
  }












  /***
 *    .d8888. db   db  .d8b.  d8888b. d88888b d8888b.      d88888b db    db d88888b d8b   db d888888b .d8888.      d888888b d888888b d88888b .88b  d88. .d8888. 
 *    88'  YP 88   88 d8' `8b 88  `8D 88'     88  `8D      88'     88    88 88'     888o  88 `~~88~~' 88'  YP        `88'   `~~88~~' 88'     88'YbdP`88 88'  YP 
 *    `8bo.   88ooo88 88ooo88 88oobY' 88ooooo 88   88      88ooooo Y8    8P 88ooooo 88V8o 88    88    `8bo.           88       88    88ooooo 88  88  88 `8bo.   
 *      `Y8b. 88~~~88 88~~~88 88`8b   88~~~~~ 88   88      88~~~~~ `8b  d8' 88~~~~~ 88 V8o88    88      `Y8b.         88       88    88~~~~~ 88  88  88   `Y8b. 
 *    db   8D 88   88 88   88 88 `88. 88.     88  .8D      88.      `8bd8'  88.     88  V888    88    db   8D        .88.      88    88.     88  88  88 db   8D 
 *    `8888Y' YP   YP YP   YP 88   YD Y88888P Y8888D'      Y88888P    YP    Y88888P VP   V8P    YP    `8888Y'      Y888888P    YP    Y88888P YP  YP  YP `8888Y' 
 *                                                                                                                                                              
 *                                                                                                                                                              
 */

  private buildAllSharedEventsItems( items: IItemDetail[] ) {

    let sharedEvents: ISharingEvent[] = [];

    //Get all events
    items.map( item => {

      if ( item.itemSharingInfo && item.itemSharingInfo.sharedEvents ) {
        item.itemSharingInfo.sharedEvents.map( event => {
          event.ServerRedirectedEmbedUrl = item.ServerRedirectedEmbedUrl;
          event.parentFolder = item.parentFolder;
          event.id = item.id;
          event.iconName = item.iconName;
          event.iconColor = item.iconColor;
          event.iconTitle = item.iconTitle;
          event.meta = item.meta;
          event.iconSearch = item.iconSearch; //Tried removing this but it caused issues with the auto-create title icons in Items.tsx so I'm adding it back.

          sharedEvents.push( event );
        });
      }
    });

    //Sort by Shared Time
    sharedEvents = sortObjectArrayByChildNumberKey( sharedEvents, 'asc', 'TimeMS' );

    return sharedEvents;

  }

  /***
 *    d88888b db    db d88888b d8b   db d888888b .d8888.      d888888b  .d8b.  d8888b. db      d88888b 
 *    88'     88    88 88'     888o  88 `~~88~~' 88'  YP      `~~88~~' d8' `8b 88  `8D 88      88'     
 *    88ooooo Y8    8P 88ooooo 88V8o 88    88    `8bo.           88    88ooo88 88oooY' 88      88ooooo 
 *    88~~~~~ `8b  d8' 88~~~~~ 88 V8o88    88      `Y8b.         88    88~~~88 88~~~b. 88      88~~~~~ 
 *    88.      `8bd8'  88.     88  V888    88    db   8D         88    88   88 88   8D 88booo. 88.     
 *    Y88888P    YP    Y88888P VP   V8P    YP    `8888Y'         YP    YP   YP Y8888P' Y88888P Y88888P 
 *                                                                                                     
 *                                                                                                     
 */

  private buildSharedEventsTable( sharedEvents: ISharingEvent[] , data: string, countToShow: number, textSearch: string, ): any {

    // let items : IItemDetail[] = itemsIn;
    let rows = [];
    let tableTitle = data;

    let priorEvent =  null;
    let filteredCount = 0;

    rows.push( <tr>
      <th></th>
      <th>Info</th>
      <th>When shared</th>
      <th>Who shared</th>
      <th>Shared with</th>
      <th>Folder</th>
      <th>File</th>
      <th>File Name</th>
    </tr> );

    //Get event rows (if visible )
    sharedEvents.map( ( event, index ) => {
      let isVisible = this.isEventVisible( textSearch, event ) === true;
      if ( rows.length < countToShow && isVisible === true ) {
          rows.push( this.createSingleEventRow( index.toFixed(0), event , priorEvent, textSearch ) );
          priorEvent = event ;
          filteredCount ++;
      } else if ( isVisible === true ) { filteredCount ++; }
    });

    return { filteredCount: filteredCount, table: this.buildTableFromRows( rows, tableTitle, itemStyles.eventsTable ) } ;

  }

  
  /***
 *    d88888b db    db d88888b d8b   db d888888b      d8888b.  .d88b.  db   d8b   db 
 *    88'     88    88 88'     888o  88 `~~88~~'      88  `8D .8P  Y8. 88   I8I   88 
 *    88ooooo Y8    8P 88ooooo 88V8o 88    88         88oobY' 88    88 88   I8I   88 
 *    88~~~~~ `8b  d8' 88~~~~~ 88 V8o88    88         88`8b   88    88 Y8   I8I   88 
 *    88.      `8bd8'  88.     88  V888    88         88 `88. `8b  d8' `8b d8'8b d8' 
 *    Y88888P    YP    Y88888P VP   V8P    YP         88   YD  `Y88P'   `8b8' `8d8'  
 *                                                                                   
 *                                                                                   
 */

private createSingleEventRow( key: string, event: ISharingEvent, priorEvent: ISharingEvent, highlight: string ) {

  let cells : any[] = [];
  cells.push( <td style={{width: '50px'}} >{ key }</td> );

  let dateStyle : React.CSSProperties = { };

  let eventTime: any = event.SharedTime.toLocaleString();
  let eventTimeTitle = event.SharedTime.toLocaleString();

  let sharedBy: any = event.sharedBy;
  let sharedWith: any = event.sharedWith;
  let FileLeafRef: any = event.FileLeafRef;
  let dateSearch = event.SharedTime.toLocaleDateString();
  let detailIcon = <td className = { itemStyles.tableIconDots }>...</td>;
  let folderIcon = <td className = { itemStyles.tableIconDots }>...</td>;
  let openItemCell = <td className = { itemStyles.tableIconDots }>...</td>;

  let isSameEvent = priorEvent !== null && event.TimeMS === priorEvent.TimeMS && event.sharedBy === priorEvent.sharedBy && event.FileLeafRef === priorEvent.FileLeafRef ? true : false;

  if ( isSameEvent !== true ) {
    folderIcon = this.buildFolderIcon( event );
    openItemCell = this.buildOpenItemCell( event, event.id.toFixed(0), `Click to preview this file` , null );
    detailIcon = this.buildDetailIcon( event, event.id.toString() );
  }

  if ( isSameEvent === true ) {
    eventTime = '...' ;
    sharedBy = eventTime = '...' ;
    FileLeafRef = eventTime = '...' ;

    //If there is highlight (search string), then highlight any text.
  } else if ( highlight && highlight.length > 0 ) {
    eventTime = getHighlightedText( eventTime, highlight ) ;
    sharedBy = getHighlightedText( sharedBy, highlight ) ;
    sharedWith = getHighlightedText( sharedWith, highlight ) ;
    FileLeafRef = getHighlightedText( FileLeafRef, highlight ) ;

  }

  cells.push( detailIcon );
  cells.push( <td style={ dateStyle } title={ eventTimeTitle } onClick = { () => this._onCTRLClickSearch(dateSearch) }>{ eventTime }</td> );
  cells.push( <td style={ null } title={ event.sharedBy } onClick = { () => this._onCTRLClickSearch(event.sharedBy) } >{ sharedBy } </td> );
  cells.push( <td style={ null } title={ null } onClick = { () => this._onCTRLClickSearch(event.sharedWith) } >{ sharedWith } </td> );

  cells.push( folderIcon );
  cells.push( openItemCell );

  cells.push( <td style={ null } title={  `Id: ${event.id} Found in folder: ${event.parentFolder}`  } onClick = { () => this._onCTRLClickSearch(event.FileLeafRef) } >{ FileLeafRef }</td> );

  let cellText: any = event.FileLeafRef;

  let cellRow = <tr style={ null }> { cells } </tr>;

  return cellRow;

}


  /***
 *    d888888b .d8888.      d88888b db    db d88888b d8b   db d888888b      db    db d888888b .d8888. d888888b d8888b. db      d88888b 
 *      `88'   88'  YP      88'     88    88 88'     888o  88 `~~88~~'      88    88   `88'   88'  YP   `88'   88  `8D 88      88'     
 *       88    `8bo.        88ooooo Y8    8P 88ooooo 88V8o 88    88         Y8    8P    88    `8bo.      88    88oooY' 88      88ooooo 
 *       88      `Y8b.      88~~~~~ `8b  d8' 88~~~~~ 88 V8o88    88         `8b  d8'    88      `Y8b.    88    88~~~b. 88      88~~~~~ 
 *      .88.   db   8D      88.      `8bd8'  88.     88  V888    88          `8bd8'    .88.   db   8D   .88.   88   8D 88booo. 88.     
 *    Y888888P `8888Y'      Y88888P    YP    Y88888P VP   V8P    YP            YP    Y888888P `8888Y' Y888888P Y8888P' Y88888P Y88888P 
 *                                                                                                                                     
 *                                                                                                                                     
 */
private isEventVisible ( textSearch: any, event: ISharingEvent ) {

  let visible = true;

  if ( textSearch.length > 0 ) {

    visible = false;

    if ( event.meta.indexOf( textSearch ) > -1 ) {
      visible = true;
    } else {
      let searchThis = getEventSearchString( event );
      if ( searchThis.toLowerCase().indexOf( textSearch.toLowerCase()) > -1 ) {
        visible = true;
      }
    }
  }
  return visible;

}


  /***
 *    d888888b d888888b d88888b .88b  d88. .d8888.      d888888b  .d8b.  d8888b. db      d88888b 
 *      `88'   `~~88~~' 88'     88'YbdP`88 88'  YP      `~~88~~' d8' `8b 88  `8D 88      88'     
 *       88       88    88ooooo 88  88  88 `8bo.           88    88ooo88 88oooY' 88      88ooooo 
 *       88       88    88~~~~~ 88  88  88   `Y8b.         88    88~~~88 88~~~b. 88      88~~~~~ 
 *      .88.      88    88.     88  88  88 db   8D         88    88   88 88   8D 88booo. 88.     
 *    Y888888P    YP    Y88888P YP  YP  YP `8888Y'         YP    YP   YP Y8888P' Y88888P Y88888P 
 *                                                                                               
 *                                                                                               
 */

  private buildItemsTable( items: IItemDetail[] | IDuplicateFile[] , itemsAreDups: boolean , data: string, countToShow: number, textSearch: string, sortKey: 'size' ): any {

    let rows = [];
    let tableTitle = data;
    let itemsSorted: any[] = [];
    let filteredCount = 0;

    rows.push( <tr>
      <th></th>
      <th></th>
      <th>Info</th>
      { this.props.itemsAreFolders === true ? 
      <th title='Number of files and folders in this folder'>#</th> :
      null  }
      <th title= { this.props.itemsAreFolders === true ? 'Total size of all files in this folder, not including subfolders' : 'File size'}>Size</th>
      <th>Modified By</th>
      <th>Modified</th>
      <th style={{paddingRight: '10px'}} title='Click Folder to go directly to the parent folder of that item.'> { this.props.itemsAreFolders === true ? 'Parent' : 'Folder'}</th>

      { this.props.itemsAreFolders !== true ? <th style={{paddingRight: '10px'}}>Version</th> : null  }
      { this.props.itemsAreFolders !== true ? <th style={{paddingRight: '10px'}}></th> : null  }
      {/* { this.props.itemsAreFolders !== true ?  : null  } */}
      <th style={{paddingRight: '10px'}}>File</th>
    </tr> );

    itemsSorted = sortObjectArrayByChildNumberKey( items, 'dec', sortKey );
    
    itemsSorted.map( ( item, index ) => {
      let isVisible = this.isVisibleItem( textSearch, item, itemsAreDups ) === true;
      if ( rows.length < countToShow && isVisible === true ) {
        rows.push( this.createSingleItemRow( index.toFixed(0), item ) );
        filteredCount ++;
      } else if ( isVisible === true ) { filteredCount ++; }
    });

    return { filteredCount: filteredCount, table: this.buildTableFromRows( rows, tableTitle, itemStyles.itemsTable ) };

  }

  
  /***
 *    d888888b d888888b d88888b .88b  d88.      d8888b.  .d88b.  db   d8b   db 
 *      `88'   `~~88~~' 88'     88'YbdP`88      88  `8D .8P  Y8. 88   I8I   88 
 *       88       88    88ooooo 88  88  88      88oobY' 88    88 88   I8I   88 
 *       88       88    88~~~~~ 88  88  88      88`8b   88    88 Y8   I8I   88 
 *      .88.      88    88.     88  88  88      88 `88. `8b  d8' `8b d8'8b d8' 
 *    Y888888P    YP    Y88888P YP  YP  YP      88   YD  `Y88P'   `8b8' `8d8'  
 *                                                                             
 *                                                                             
 */

private createSingleItemRow( key: string, item: IItemDetail ) {

  let itemFolder : any = this.props.itemsAreFolders === true ? item : null;
  let folder: IFolderDetail = itemFolder;

  let created = new Date( item.created );
  let modified = new Date( item.modified );

  let cells : any[] = [];
  cells.push( <td style={{width: '50px'}} >{ key }</td> );

  let id = this.props.itemsAreDups === true ? item.id.toString() : item.id.toString()  ;
  let detailItemIcon = this.buildDetailIcon( item, id );
  let detailMediaIcon = this.buildMediaIconCell( item, id );
  
  let userStyle: any =  { width: null } ;
  let userTitle = null;

  if ( item.authorTitle !== item.editorTitle ) { 
    userStyle.color = 'red';
    userStyle.fontWeight = 600;
    userTitle = `Edited by ${ item.editorTitle }`;
  }

  cells.push( detailMediaIcon );
  cells.push( detailItemIcon );

  if ( this.props.itemsAreFolders === true ) {
    cells.push( <td style={{width: null }} >{ getCountLabel( folder.directCount, 0 ) }</td> );
    // console.log('getting sizeLabel: ', folder );
    cells.push( <td style={{ paddingRight: '15px' }} >{ getSizeLabel( folder.directSize ) }</td> );

  } else {
    cells.push( <td style={{ paddingRight: '15px' }} >{ getSizeLabel( item.size ) }</td> );
  }

  cells.push( <td style={ userStyle } title={ userTitle }>{ item.authorTitle }</td> );
  let dateStyle : React.CSSProperties = { };
  let dateTitle : string = '';

  if ( item.createMs < item.modMs ) {
    dateStyle.color = 'blue';
    dateTitle = `Modified: ${ modified.toLocaleString() }`;
  
  } else if ( item.createMs > item.modMs ) {
    dateStyle.color = 'red';
    dateStyle.fontWeight = 600;
    dateTitle = `Modified Before Created!!! : ${ modified.toLocaleString() }`;

  }

  cells.push( <td style={dateStyle} title={ dateTitle }>{ created.toLocaleString([], { year: 'numeric', month: 'numeric', day: 'numeric', hour: '2-digit', minute: '2-digit' }) }</td> );

  if ( this.props.itemsAreDups !== true ) {
    cells.push( this.buildFolderIcon( item ) );
  }

  // cells.push( <td style={cellMaxStyle}><a href={ item.FileRef } target={ '_blank' }>{ item.FileLeafRef }</a></td> );
  if ( this.props.itemsAreFolders === false ) {

    cells.push( <td style={{paddingRight: '15px' }} >{ item.version.string }</td> );

    if ( item.checkedOutCurrentUser === true ) {
      cells.push( <td style={ null } >{ buildClickableIcon('eXTremeStorage', StdIcons.CheckedOutByYou , `You checked out this item.  Your Id is: ${ item.checkedOutId }`, '#a4262c', this._onClickDataSearch.bind(this), null 'CheckedOutToYou' ) }</td> );

    } else if ( item.checkedOutId ) {
      cells.push( <td style={ null } >{ buildClickableIcon('eXTremeStorage', StdIcons.CheckedOutByOther , `Checked out by: ${ item.checkedOutId }`, 'black', this._onClickDataSearch.bind(this), null 'CheckedOutToYou') }</td> );
      
    } else { cells.push( <td></td> ); }

  }

  let cellText: any = item.FileLeafRef;
  //For duplicate files, this will show the relative path.
  //BUT to help when there are deep folders, it will show based on the common parent folder, not full folder url because it can be to long
  if ( this.props.itemsAreDups === true ) {
    cellText = <span><span style={{ fontWeight: 600 }}>{'../' + this.commonParent }</span><span>{ item.parentFolder.replace( this.commonRelativePath, '' ) } </span></span> ; //commonParent
  } 
  cells.push( this.buildOpenItemCell( item, item.id.toFixed(0) , cellText, cellText ) );

  let cellRow = <tr style={ null }> { cells } </tr>;

  return cellRow;

}


  /***
 *    d888888b .d8888.      d888888b d888888b d88888b .88b  d88.      db    db d888888b .d8888. d888888b d8888b. db      d88888b 
 *      `88'   88'  YP        `88'   `~~88~~' 88'     88'YbdP`88      88    88   `88'   88'  YP   `88'   88  `8D 88      88'     
 *       88    `8bo.           88       88    88ooooo 88  88  88      Y8    8P    88    `8bo.      88    88oooY' 88      88ooooo 
 *       88      `Y8b.         88       88    88~~~~~ 88  88  88      `8b  d8'    88      `Y8b.    88    88~~~b. 88      88~~~~~ 
 *      .88.   db   8D        .88.      88    88.     88  88  88       `8bd8'    .88.   db   8D   .88.   88   8D 88booo. 88.     
 *    Y888888P `8888Y'      Y888888P    YP    Y88888P YP  YP  YP         YP    Y888888P `8888Y' Y888888P Y8888P' Y88888P Y88888P 
 *                                                                                                                               
 *                                                                                                                               
 */
private isVisibleItem ( textSearch: string, item: IItemDetail, itemsAreDups: boolean ) {

  let visible : boolean = true;
  let anyTextSearch : any = textSearch;

  if ( textSearch.length > 0 ) {
    visible = false;

    if ( item.meta.length > 0 ) {
      if ( item.meta.indexOf( anyTextSearch ) > -1 ) {
        visible = true;
      }
    }

    if ( visible === false ) {
      let searchThis = getItemSearchString( item, itemsAreDups, false );
      if ( searchThis.toLowerCase().indexOf( textSearch.toLowerCase()) > -1 ) {
        visible = true;
      }
    }
  }

  return visible;

}


  /***
 *    d8888b. db    db d8888b. .d8888.      d888888b  .d8b.  d8888b. db      d88888b 
 *    88  `8D 88    88 88  `8D 88'  YP      `~~88~~' d8' `8b 88  `8D 88      88'     
 *    88   88 88    88 88oodD' `8bo.           88    88ooo88 88oooY' 88      88ooooo 
 *    88   88 88    88 88~~~     `Y8b.         88    88~~~88 88~~~b. 88      88~~~~~ 
 *    88  .8D 88b  d88 88      db   8D         88    88   88 88   8D 88booo. 88.     
 *    Y8888D' ~Y8888P' 88      `8888Y'         YP    YP   YP Y8888P' Y88888P Y88888P 
 *                                                                                   
 *                                                                                   
 */
  /**
   * Same as buildItemsTable but only for when  } objectType === 'Duplicates'
   * @param items
   * @param itemsAreDups 
   * @param objectType 
   * @param data 
   * @param countToShow 
   * @param textSearch 
   * @param sortKey 
   */
  private buildDupsTable( items: IDuplicateFile[] , itemsAreDups: boolean , data: string, countToShow: number, textSearch: string, sortKey: 'size' ): any {

    let rows = [];
    let tableTitle = data;
    let itemsSorted: any[] = [];
    let filteredCount = 0;
    
    rows.push( <tr>
      <th></th>
      <th title='Number of files with the same exact name and extension'>#</th>
      <th>Size</th>
      <th style={{paddingRight: '10px'}}>File</th>
    </tr> );

    itemsSorted = sortObjectArrayByChildNumberKey( items, 'dec', sortKey );
    
    itemsSorted.map( ( item, index ) => {
      let isVisible = this.isVisibleItem( textSearch, item, itemsAreDups ) === true;
      if ( rows.length < countToShow && isVisible === true ) {
        rows.push( this.createSingleDupRow( index.toFixed(0), item ) );
        filteredCount ++;
      } else if ( isVisible === true ) { filteredCount ++; }
    });

    return { filteredCount: filteredCount, table: this.buildTableFromRows( rows, tableTitle, itemStyles.itemsTable ) };

  }
  

  /***
 *    d8888b. db    db d8888b.      d8888b.  .d88b.  db   d8b   db 
 *    88  `8D 88    88 88  `8D      88  `8D .8P  Y8. 88   I8I   88 
 *    88   88 88    88 88oodD'      88oobY' 88    88 88   I8I   88 
 *    88   88 88    88 88~~~        88`8b   88    88 Y8   I8I   88 
 *    88  .8D 88b  d88 88           88 `88. `8b  d8' `8b d8'8b d8' 
 *    Y8888D' ~Y8888P' 88           88   YD  `Y88P'   `8b8' `8d8'  
 *                                                                 
 *                                                                 
 */
  private createSingleDupRow( key: string, item: IDuplicateFile ) {

    // let created = new Date( item.created );
    let detailIcon = 'DocumentSearch';
    let detailIconStyle = 'black';

    let cells : any[] = [];
    cells.push( <td style={{width: '50px', textAlign: 'center' }} >{ key }</td> );

    cells.push( <td style={{width: '50px', textAlign: 'center' }} >{ item.summary.count }</td> );
    cells.push( <td style={{width: '100px', textAlign: 'center' }} >{ item.summary.sizeLabel }</td> );

    const iconStyles: any = { root: {
      fontSize: 'larger',
      color: item.iconColor,
      padding: '0px 4px 0px 10px',
    }};

    cells.push( this.buildOpenItemCell( item, item.FileLeafRef, item.FileLeafRef, item.FileLeafRef ) );
  
    let cellRow = <tr> { cells } </tr>;

    return cellRow;
  
  }


  /***
 *     .d88b.  d8888b. d88888b d8b   db      d888888b d888888b d88888b .88b  d88.       .o88b. d88888b db      db      
 *    .8P  Y8. 88  `8D 88'     888o  88        `88'   `~~88~~' 88'     88'YbdP`88      d8P  Y8 88'     88      88      
 *    88    88 88oodD' 88ooooo 88V8o 88         88       88    88ooooo 88  88  88      8P      88ooooo 88      88      
 *    88    88 88~~~   88~~~~~ 88 V8o88         88       88    88~~~~~ 88  88  88      8b      88~~~~~ 88      88      
 *    `8b  d8' 88      88.     88  V888        .88.      88    88.     88  88  88      Y8b  d8 88.     88booo. 88booo. 
 *     `Y88P'  88      Y88888P VP   V8P      Y888888P    YP    Y88888P YP  YP  YP       `Y88P' Y88888P Y88888P Y88888P 
 *                                                                                                                     
 *                                                                                                                     
 */
  private buildOpenItemCell ( item: IItemDetail | IDuplicateFile | ISharingEvent, itemId: string, titleName: any, text: any ) {
    let cell = <td className={itemStyles.cellMaxStyle} onClick={ this._onClickItem.bind(this)} 
    id={ itemId } 
    title={ `Item ID: ${ itemId } Item Name: ${ titleName }` }
  >
    { <Icon iconName= { item.iconName } data-search={ item.iconSearch } style={ { fontSize: 'larger', color: item.iconColor, padding: '0px 15px 0px 0px', } }></Icon> }
    { text }</td>;

    return cell;
  }


  /***
 *     .d88b.  d8b   db       .o88b. db      d888888b  .o88b. db   dD      d888888b d888888b d88888b .88b  d88. 
 *    .8P  Y8. 888o  88      d8P  Y8 88        `88'   d8P  Y8 88 ,8P'        `88'   `~~88~~' 88'     88'YbdP`88 
 *    88    88 88V8o 88      8P      88         88    8P      88,8P           88       88    88ooooo 88  88  88 
 *    88    88 88 V8o88      8b      88         88    8b      88`8b           88       88    88~~~~~ 88  88  88 
 *    `8b  d8' 88  V888      Y8b  d8 88booo.   .88.   Y8b  d8 88 `88.        .88.      88    88.     88  88  88 
 *     `Y88P'  VP   V8P       `Y88P' Y88888P Y888888P  `Y88P' YP   YD      Y888888P    YP    Y88888P YP  YP  YP 
 *                                                                                                              
 *                                                                                                              
 */
private _onClickItem( event ) {
  // console.log( event );
  console.log( event.currentTarget.id );

  if ( this.props.itemType === 'Items' || this.props.itemType === 'Shared' ) {
    let clickThisItem = parseInt(event.currentTarget.id) ;
    let items: IItemDetail[] = this.itemsAny;
    let selectedItem = null;

    items.map( item => {
      let openThisLink =  item.ServerRedirectedEmbedUrl;
      if ( !openThisLink || openThisLink.length === 0 ) { 
        openThisLink = item.FileRef ;
       }

      if ( item.id === clickThisItem ) { 
        if ( !item.ServerRedirectedEmbedUrl ) {
          window.open( openThisLink, "_blank");
        }
        selectedItem = item;
       }
    });

    this.setState({ 
      selectedItem: selectedItem,
      showPreview: selectedItem && selectedItem.ServerRedirectedEmbedUrl ? true : false,
      showItems: [],  //Clear any duplicate items
    });

  } else if ( this.props.itemType === 'Duplicates' ) {
    let clickThisItem = event.currentTarget.id ; //For some reason this has the DupName but then is set to undefined after the next line.
    let duplicates: IDuplicateFile[] = this.props.duplicateInfo.duplicates ;
    let showItems : IItemDetail[] = [];
    duplicates.map( dup => {
      if ( dup.FileLeafRef === event.currentTarget.id ) {
        showItems = dup.items;
      }
    });

    this.setState({ 
      selectedItem: null,
      showPreview: null,
      showItems: showItems,
    });

  } else { alert('Ooops!  We haven\`t made it so you can click on an event yet :( ') ; }
}



  /***
 *    d8888b. d88888b d888888b  .d8b.  d888888b db           d888888b  .o88b.  .d88b.  d8b   db 
 *    88  `8D 88'     `~~88~~' d8' `8b   `88'   88             `88'   d8P  Y8 .8P  Y8. 888o  88 
 *    88   88 88ooooo    88    88ooo88    88    88              88    8P      88    88 88V8o 88 
 *    88   88 88~~~~~    88    88~~~88    88    88              88    8b      88    88 88 V8o88 
 *    88  .8D 88.        88    88   88   .88.   88booo.        .88.   Y8b  d8 `8b  d8' 88  V888 
 *    Y8888D' Y88888P    YP    YP   YP Y888888P Y88888P      Y888888P  `Y88P'  `Y88P'  VP   V8P 
 *                                                                                              
 *                                                                                              
 */
  private buildDetailIcon ( itemIn: IItemDetail | ISharingEvent , id: string ) {

    let item: any = itemIn; //Needed to use same component with different interfaces that may not match
    let iconSearch : IKnownMeta = item.iconSearch;

    let detailIcon = item.isMedia === true ? fpsAppIcons.ImageSearchRed : fpsAppIcons.DocumentSearch;

    if ( item.itemSharingInfo ) { 
      detailIcon = fpsAppIcons.SharedItem;
      iconSearch = 'WasShared';

    } else if ( item.uniquePerms === true ) { 
      detailIcon = fpsAppIcons.UniquePerms;
      iconSearch = 'UniquePermissions';

    }

    let iconCell = <td className = { itemStyles.tableIcons } style={{width: '50px', cursor: 'pointer', position: 'relative' }} 
      onClick={ this._onClickItemDetail.bind(this)} id={ id } data-search = { iconSearch }
      title={ `See all Item Details.` }>
      <div style={{ position: 'relative' }} onClick={ this._onClickItemDetail.bind(this)} id={ id } data-search = { iconSearch }>{ detailIcon }</div>
    </td>;

    return iconCell;

  }

  
/***
 *     .d88b.  d8b   db       .o88b. db      d888888b  .o88b. db   dD      d888888b d888888b d88888b .88b  d88.      d8888b. d88888b d888888b  .d8b.  d888888b db      
 *    .8P  Y8. 888o  88      d8P  Y8 88        `88'   d8P  Y8 88 ,8P'        `88'   `~~88~~' 88'     88'YbdP`88      88  `8D 88'     `~~88~~' d8' `8b   `88'   88      
 *    88    88 88V8o 88      8P      88         88    8P      88,8P           88       88    88ooooo 88  88  88      88   88 88ooooo    88    88ooo88    88    88      
 *    88    88 88 V8o88      8b      88         88    8b      88`8b           88       88    88~~~~~ 88  88  88      88   88 88~~~~~    88    88~~~88    88    88      
 *    `8b  d8' 88  V888      Y8b  d8 88booo.   .88.   Y8b  d8 88 `88.        .88.      88    88.     88  88  88      88  .8D 88.        88    88   88   .88.   88booo. 
 *     `Y88P'  VP   V8P       `Y88P' Y88888P Y888888P  `Y88P' YP   YD      Y888888P    YP    Y88888P YP  YP  YP      Y8888D' Y88888P    YP    YP   YP Y888888P Y88888P 
 *                                                                                                                                                                     
 *                                                                                                                                                                     
 */
private _onClickItemDetail( event ){
  console.log( event );
  console.log( event.currentTarget.id );
  let showThisType = event.currentTarget.id;

  let selectedItem = null;

  if ( event.ctrlKey === true ) {
    let fileType = event.currentTarget.dataset.search;
    //Apply file type filter only on CTRL-Click, else other filter
    if ( fileType === this.state.textSearch ) { fileType = ''; }
    this._searchForItems ( fileType );

  } else {
    if ( this.props.itemType === 'Items' || this.props.itemType === 'Shared' ) {
      this.itemsAny.map( item => {
        let checkThis = this.props.itemsAreDups === true ? item.id : item.id  ;
        if ( checkThis == showThisType ) { selectedItem = item ; }
      });
      this.setState({
        showItem: true,
        showPreview: false,
        selectedItem: selectedItem,
        showItems: [],
      });

    } else {
      console.log('WHOOOPPS... THIS SHOULD NO HAVE HAPPEND - EsItems.tsx ~654');
    }
  }
}




  /***
 *    d88888b  .d88b.  db      d8888b. d88888b d8888b.      d888888b  .o88b.  .d88b.  d8b   db 
 *    88'     .8P  Y8. 88      88  `8D 88'     88  `8D        `88'   d8P  Y8 .8P  Y8. 888o  88 
 *    88ooo   88    88 88      88   88 88ooooo 88oobY'         88    8P      88    88 88V8o 88 
 *    88~~~   88    88 88      88   88 88~~~~~ 88`8b           88    8b      88    88 88 V8o88 
 *    88      `8b  d8' 88booo. 88  .8D 88.     88 `88.        .88.   Y8b  d8 `8b  d8' 88  V888 
 *    YP       `Y88P'  Y88888P Y8888D' Y88888P 88   YD      Y888888P  `Y88P'  `Y88P'  VP   V8P 
 *                                                                                             
 *                                                                                             
 */

private buildFolderIcon ( itemIn: IItemDetail | ISharingEvent ) {

  let item: any = itemIn; //Added any type so that itemId can be found on either type

  let itemId = item.id ? item.id : item.itemId;
  let folderIcon = buildAppWarnIcon('eXTremeStorage', StdIcons.FabricMovetoFolder, `Go to folder: ${ item.parentFolder }`, 'black');
  let iconCell = <td className = { itemStyles.folderIcons } 
    onClick={ this._onClickFolder.bind(this)} id={ itemId.toFixed(0) }
    title={ `Go to parent folder: ${ item.parentFolder }`} >
    {/* { <Icon iconName= {'FabricMovetoFolder'} style={{ padding: '4px 4px', fontSize: 'large' }}></Icon> } */}
    { folderIcon }
  </td>;
  return iconCell;

}


  /***
 *     .d88b.  d8b   db       .o88b. db      d888888b  .o88b. db   dD      d88888b  .d88b.  db      d8888b. d88888b d8888b. 
 *    .8P  Y8. 888o  88      d8P  Y8 88        `88'   d8P  Y8 88 ,8P'      88'     .8P  Y8. 88      88  `8D 88'     88  `8D 
 *    88    88 88V8o 88      8P      88         88    8P      88,8P        88ooo   88    88 88      88   88 88ooooo 88oobY' 
 *    88    88 88 V8o88      8b      88         88    8b      88`8b        88~~~   88    88 88      88   88 88~~~~~ 88`8b   
 *    `8b  d8' 88  V888      Y8b  d8 88booo.   .88.   Y8b  d8 88 `88.      88      `8b  d8' 88booo. 88  .8D 88.     88 `88. 
 *     `Y88P'  VP   V8P       `Y88P' Y88888P Y888888P  `Y88P' YP   YD      YP       `Y88P'  Y88888P Y8888D' Y88888P 88   YD 
 *                                                                                                                          
 *                                                                                                                          
 */
private _onClickFolder( event ) {
  // console.log( event );
  console.log( event.currentTarget.id );
  let clickThisItem = parseInt(event.currentTarget.id);

  this.itemsAny.map( item => {
    if ( item.id === clickThisItem ) { 
      window.open( item.parentFolder, "_blank");
    }
  });
}

  /***
 *    .88b  d88. d88888b d8888b. d888888b  .d8b.        .o88b. d88888b db      db      
 *    88'YbdP`88 88'     88  `8D   `88'   d8' `8b      d8P  Y8 88'     88      88      
 *    88  88  88 88ooooo 88   88    88    88ooo88      8P      88ooooo 88      88      
 *    88  88  88 88~~~~~ 88   88    88    88~~~88      8b      88~~~~~ 88      88      
 *    88  88  88 88.     88  .8D   .88.   88   88      Y8b  d8 88.     88booo. 88booo. 
 *    YP  YP  YP Y88888P Y8888D' Y888888P YP   YP       `Y88P' Y88888P Y88888P Y88888P 
 *                                                                                     
 *                                                                                     
 */
  private buildMediaIconCell ( item: IItemDetail, id: string ) {
    
    let MediaIcons: any[] = item.isMedia ? this.buildMediaIcons( item ) : [];

    let iconCell = <td className = { itemStyles.tableIcons } style={{width: MediaIcons.length > 0 ? '50px' : '0px', position: 'relative' }}>
      <div style={{ display: 'inline-block', position: 'absolute', marginLeft: '3px', top: '0px' }}> { MediaIcons } </div>
    </td>;

    return iconCell;

  }


  /***
 *    .88b  d88. d88888b d8888b. d888888b  .d8b.       d888888b  .o88b.  .d88b.  d8b   db .d8888. 
 *    88'YbdP`88 88'     88  `8D   `88'   d8' `8b        `88'   d8P  Y8 .8P  Y8. 888o  88 88'  YP 
 *    88  88  88 88ooooo 88   88    88    88ooo88         88    8P      88    88 88V8o 88 `8bo.   
 *    88  88  88 88~~~~~ 88   88    88    88~~~88         88    8b      88    88 88 V8o88   `Y8b. 
 *    88  88  88 88.     88  .8D   .88.   88   88        .88.   Y8b  d8 `8b  d8' 88  V888 db   8D 
 *    YP  YP  YP Y88888P Y8888D' Y888888P YP   YP      Y888888P  `Y88P'  `Y88P'  VP   V8P `8888Y' 
 *                                                                                                
 *                                                                                                
 */
  private buildMediaIcons( item: any ) { //IItemDetail
    let MediaIcons = [];
    if ( item.isMedia ) {

      if ( item.MediaServiceOCR ) {
        MediaIcons.push(  <Icon iconName= { 'CircleShapeSolid' } style={{ top: '2px', left: '2px', fontSize: '6px', position: 'absolute', color: 'dimgray', cursor: 'mouse' }} 
          onClick={ this._onClickItemDetail.bind(this)} id={ 'MediaServiceOCR' } data-search = { 'MediaServiceOCR' } title='MediaServiceOCR'></Icon> );
      }
      if ( item.MediaServiceAutoTags ) {
        MediaIcons.push(  <Icon iconName= { 'TagSolid' } style={{ top: '1px', left: '12px', fontSize: '9px', position: 'absolute', color: 'dimgray', cursor: 'mouse' }} 
          onClick={ this._onClickItemDetail.bind(this)} id={ 'MediaServiceAutoTags' } data-search = { 'MediaServiceAutoTags' } title='MediaServiceAutoTags'></Icon> );
      }
      if ( item.MediaServiceKeyPoints ) {
        MediaIcons.push(  <Icon iconName= { 'Location' } style={{ top: '10px', left: '2px', fontSize: '5px', position: 'absolute', color: 'dimgray', cursor: 'mouse' }} 
          onClick={ this._onClickItemDetail.bind(this)} id={ 'MediaServiceKeyPoints' } data-search = { 'MediaServiceKeyPoints' } title='MediaServiceKeyPoints'></Icon> );
      }
      if ( item.MediaServiceLocation ) {
        MediaIcons.push(  <Icon iconName= { 'POISolid' } style={{ top: '11px', left: '12px', fontSize: '8px', position: 'absolute', color: 'dimgray', cursor: 'mouse' }} 
          onClick={ this._onClickItemDetail.bind(this)} id={ 'MediaServiceLocation' } data-search = { 'MediaServiceLocation' } title='MediaServiceLocation'></Icon> );
      }

    }

    return MediaIcons;

  }
  


  /***
 *    .d8888. d88888b  .d8b.  d8888b.  .o88b. db   db      d8888b.  .d88b.  db    db 
 *    88'  YP 88'     d8' `8b 88  `8D d8P  Y8 88   88      88  `8D .8P  Y8. `8b  d8' 
 *    `8bo.   88ooooo 88ooo88 88oobY' 8P      88ooo88      88oooY' 88    88  `8bd8'  
 *      `Y8b. 88~~~~~ 88~~~88 88`8b   8b      88~~~88      88~~~b. 88    88  .dPYb.  
 *    db   8D 88.     88   88 88 `88. Y8b  d8 88   88      88   8D `8b  d8' .8P  Y8. 
 *    `8888Y' Y88888P YP   YP 88   YD  `Y88P' YP   YP      Y8888P'  `Y88P'  YP    YP 
 *                                                                                   
 *                                                                                   
 */

private buildSearchBox( testSearch: string ) {
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
      value={ testSearch }

    />
    <div className={styles.searchStatus}>
      { `Search all ${ getCommaSepLabel( this.itemsAny.length) } items [ ${ getSizeLabel( this.state.totalSize ) } ]` }
      { /* 'Searching ' + (this.state.searchType !== 'all' ? this.state.filteredTiles.length : ' all' ) + ' items' */ }
    </div>
  </div>;

  return searchBox;

}
  

  /***
 *     .d88b.  d8b   db       .o88b. db      d888888b  .o88b. db   dD      d888888b db    db d8888b. d88888b 
 *    .8P  Y8. 888o  88      d8P  Y8 88        `88'   d8P  Y8 88 ,8P'      `~~88~~' `8b  d8' 88  `8D 88'     
 *    88    88 88V8o 88      8P      88         88    8P      88,8P           88     `8bd8'  88oodD' 88ooooo 
 *    88    88 88 V8o88      8b      88         88    8b      88`8b           88       88    88~~~   88~~~~~ 
 *    `8b  d8' 88  V888      Y8b  d8 88booo.   .88.   Y8b  d8 88 `88.         88       88    88      88.     
 *     `Y88P'  VP   V8P       `Y88P' Y88888P Y888888P  `Y88P' YP   YD         YP       YP    88      Y88888P 
 *                                                                                                           
 *                                                                                                           
 */
  /**
   * This is the same as _onCTRLClickSearch except it was clicked via the file type icon at the top.
   * @param event 
   */
  private _onClickType( event ) {
    console.log( '_onClickType:',  event );
    let iconSearch = event.currentTarget.id;
    let fileType = event.currentTarget.dataset.search;
    //This is a quick "clear search" feature
    if ( iconSearch === this.state.textSearch ) { iconSearch = ''; }
    this._searchForItems ( iconSearch );

  }

/***
 *     .d88b.  d8b   db       .o88b. db      d888888b  .o88b. db   dD      d8888b.  .d8b.  d888888b  .d8b.       .d8888. d88888b  .d8b.  d8888b.  .o88b. db   db 
 *    .8P  Y8. 888o  88      d8P  Y8 88        `88'   d8P  Y8 88 ,8P'      88  `8D d8' `8b `~~88~~' d8' `8b      88'  YP 88'     d8' `8b 88  `8D d8P  Y8 88   88 
 *    88    88 88V8o 88      8P      88         88    8P      88,8P        88   88 88ooo88    88    88ooo88      `8bo.   88ooooo 88ooo88 88oobY' 8P      88ooo88 
 *    88    88 88 V8o88      8b      88         88    8b      88`8b        88   88 88~~~88    88    88~~~88        `Y8b. 88~~~~~ 88~~~88 88`8b   8b      88~~~88 
 *    `8b  d8' 88  V888      Y8b  d8 88booo.   .88.   Y8b  d8 88 `88.      88  .8D 88   88    88    88   88      db   8D 88.     88   88 88 `88. Y8b  d8 88   88 
 *     `Y88P'  VP   V8P       `Y88P' Y88888P Y888888P  `Y88P' YP   YD      Y8888D' YP   YP    YP    YP   YP      `8888Y' Y88888P YP   YP 88   YD  `Y88P' YP   YP 
 *                                                                                                                                                               
 *                                                                                                                                                               
 */
  /**
   * This is the same as _onCTRLClickSearch except it was clicked via the file type icon at the top.
   * @param event 
   */
  private _onClickDataSearch( event ) {
    console.log( '_onClickType:',  event );
    let textSearch = event.currentTarget.dataset.search;
    //This is a quick "clear search" feature
    if ( textSearch === this.state.textSearch ) { textSearch = ''; }
    this._searchForItems ( textSearch );

  }

  
  /***
 *     .d88b.  d8b   db       .o88b. db      d888888b  .o88b. db   dD      .d8888. d88888b  .d8b.  d8888b.  .o88b. db   db 
 *    .8P  Y8. 888o  88      d8P  Y8 88        `88'   d8P  Y8 88 ,8P'      88'  YP 88'     d8' `8b 88  `8D d8P  Y8 88   88 
 *    88    88 88V8o 88      8P      88         88    8P      88,8P        `8bo.   88ooooo 88ooo88 88oobY' 8P      88ooo88 
 *    88    88 88 V8o88      8b      88         88    8b      88`8b          `Y8b. 88~~~~~ 88~~~88 88`8b   8b      88~~~88 
 *    `8b  d8' 88  V888      Y8b  d8 88booo.   .88.   Y8b  d8 88 `88.      db   8D 88.     88   88 88 `88. Y8b  d8 88   88 
 *     `Y88P'  VP   V8P       `Y88P' Y88888P Y888888P  `Y88P' YP   YD      `8888Y' Y88888P YP   YP 88   YD  `Y88P' YP   YP 
 *                                                                                                                         
 *                                                                                                                         
 */
  /**
   * This is used for sharingEvents originally when clicking on an item.
   * Originally no CTRL press is required but called it that in case I use it elsewhere on normal items where 
   *    CTRL press might be required to distinguish a normal existing click
   * 
   * @param searchThis 
   */
  private _onCTRLClickSearch( searchThis: string ) : void {
    //This is a quick "clear search" feature

    if ( searchThis === this.state.textSearch ) { searchThis = ''; }
    this._searchForItems ( searchThis );

  }


  /***
 *    .d8888. d88888b  .d8b.  d8888b.  .o88b. db   db      d88888b  .d88b.  d8888b.      d888888b d888888b d88888b .88b  d88. .d8888. 
 *    88'  YP 88'     d8' `8b 88  `8D d8P  Y8 88   88      88'     .8P  Y8. 88  `8D        `88'   `~~88~~' 88'     88'YbdP`88 88'  YP 
 *    `8bo.   88ooooo 88ooo88 88oobY' 8P      88ooo88      88ooo   88    88 88oobY'         88       88    88ooooo 88  88  88 `8bo.   
 *      `Y8b. 88~~~~~ 88~~~88 88`8b   8b      88~~~88      88~~~   88    88 88`8b           88       88    88~~~~~ 88  88  88   `Y8b. 
 *    db   8D 88.     88   88 88 `88. Y8b  d8 88   88      88      `8b  d8' 88 `88.        .88.      88    88.     88  88  88 db   8D 
 *    `8888Y' Y88888P YP   YP 88   YD  `Y88P' YP   YP      YP       `Y88P'  88   YD      Y888888P    YP    Y88888P YP  YP  YP `8888Y' 
 *                                                                                                                                    
 *                                                                                                                                    
 */

public _searchForItems = (item): void => {
  //This sends back the correct pivot category which matches the category on the tile.
  let e: any = event;
  console.log('searchForItems: e',e);
  console.log('searchForItems: item', item);
  console.log('searchForItems: this', this);

  this.setState({ textSearch: item });
}



  /***
 *    d888888b db    db d8888b. d88888b      .d8888. db      d888888b d8888b. d88888b d8888b. 
 *    `~~88~~' `8b  d8' 88  `8D 88'          88'  YP 88        `88'   88  `8D 88'     88  `8D 
 *       88     `8bd8'  88oodD' 88ooooo      `8bo.   88         88    88   88 88ooooo 88oobY' 
 *       88       88    88~~~   88~~~~~        `Y8b. 88         88    88   88 88~~~~~ 88`8b   
 *       88       88    88      88.          db   8D 88booo.   .88.   88  .8D 88.     88 `88. 
 *       YP       YP    88      Y88888P      `8888Y' Y88888P Y888888P Y8888D' Y88888P 88   YD 
 *                                                                                            
 *                                                                                            
 */
  private _typeSlider(newValue: number){
    this.setState({
      rankSlider: newValue,
    });
  }


/***
*    d8888b. db    db d888888b db      d8888b.      d888888b  .d8b.  d8888b. db      d88888b 
*    88  `8D 88    88   `88'   88      88  `8D      `~~88~~' d8' `8b 88  `8D 88      88'     
*    88oooY' 88    88    88    88      88   88         88    88ooo88 88oooY' 88      88ooooo 
*    88~~~b. 88    88    88    88      88   88         88    88~~~88 88~~~b. 88      88~~~~~ 
*    88   8D 88b  d88   .88.   88booo. 88  .8D         88    88   88 88   8D 88booo. 88.     
*    Y8888P' ~Y8888P' Y888888P Y88888P Y8888D'         YP    YP   YP Y8888P' Y88888P Y88888P 
*                                                                                            
*                                                                                            
*/

private buildTableFromRows( rows, tableTitle, tableClassName ) {
  
  let table = <div style={{marginRight: '10px'}} className = { tableClassName }>
    <h3 style={{ textAlign: 'center' }}> { tableTitle }</h3>
    {/* <table style={{padding: '0 20px'}}> */}
    <table style={{  }} id="Select-b">
      { rows }
    </table>
  </div>;
  return table;

}


  /***
 *     .d88b.  d8b   db      d8888b. d888888b  .d8b.  db       .d88b.   d888b       d8888b. d888888b .d8888. .88b  d88. d888888b .d8888. .d8888. 
 *    .8P  Y8. 888o  88      88  `8D   `88'   d8' `8b 88      .8P  Y8. 88' Y8b      88  `8D   `88'   88'  YP 88'YbdP`88   `88'   88'  YP 88'  YP 
 *    88    88 88V8o 88      88   88    88    88ooo88 88      88    88 88           88   88    88    `8bo.   88  88  88    88    `8bo.   `8bo.   
 *    88    88 88 V8o88      88   88    88    88~~~88 88      88    88 88  ooo      88   88    88      `Y8b. 88  88  88    88      `Y8b.   `Y8b. 
 *    `8b  d8' 88  V888      88  .8D   .88.   88   88 88booo. `8b  d8' 88. ~8~      88  .8D   .88.   db   8D 88  88  88   .88.   db   8D db   8D 
 *     `Y88P'  VP   V8P      Y8888D' Y888888P YP   YP Y88888P  `Y88P'   Y888P       Y8888D' Y888888P `8888Y' YP  YP  YP Y888888P `8888Y' `8888Y' 
 *                                                                                                                                               
 *                                                                                                                                               
 */
  private _onDialogDismiss( event ) {
    this.setState({
      showItem: false,
      showPreview: false,
      selectedItem: null,
      showItems: [],
    });
  }

  /***
 *     .d88b.  d8b   db       .o88b. db       .d88b.  .d8888. d88888b      d8888b. d88888b d888888b  .d8b.  d888888b db      
 *    .8P  Y8. 888o  88      d8P  Y8 88      .8P  Y8. 88'  YP 88'          88  `8D 88'     `~~88~~' d8' `8b   `88'   88      
 *    88    88 88V8o 88      8P      88      88    88 `8bo.   88ooooo      88   88 88ooooo    88    88ooo88    88    88      
 *    88    88 88 V8o88      8b      88      88    88   `Y8b. 88~~~~~      88   88 88~~~~~    88    88~~~88    88    88      
 *    `8b  d8' 88  V888      Y8b  d8 88booo. `8b  d8' db   8D 88.          88  .8D 88.        88    88   88   .88.   88booo. 
 *     `Y88P'  VP   V8P       `Y88P' Y88888P  `Y88P'  `8888Y' Y88888P      Y8888D' Y88888P    YP    YP   YP Y888888P Y88888P 
 *                                                                                                                           
 *                                                                                                                           
 */
  private _onCloseItemDetail( event ){
    this.setState({
      showItem: false,
      showPreview: false,
      selectedItem: null,
      showItems: [],
    });
  }

  private _onPreviewClick( event ){

    return;
    this.setState({
      showPreview: false,
      selectedItem: null,
    });
  }

}
