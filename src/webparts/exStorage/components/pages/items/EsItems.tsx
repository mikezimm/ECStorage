import * as React from 'react';
import styles from '../../ExStorage.module.scss';
import itemStyles from './Items.module.scss';

import { IEsItemsProps } from './IEsItemsProps';
import { IEsItemsState } from './IEsItemsState';
import { IExStorageState, IEXStorageList, IEXStorageBatch, IBatchData, IUserSummary, IFileType, IItemDetail, IDuplicateFile, IItemType, IFolderDetail } from '../../IExStorageState';
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
import { getHelpfullErrorV2 } from '@mikezimm/npmfunctions/dist/Services/Logging/ErrorHandler';

// import { sortObjectArrayByChildNumberKey, sortNumberArray } from '@mikezimm/npmfunctions/dist/Services/Arrays/sorting';
import { sortObjectArrayByChildNumberKey, } from '@mikezimm/npmfunctions/dist/Services/Arrays/sorting';

import * as fpsAppIcons from '@mikezimm/npmfunctions/dist/Icons/standardExStorage';

import { createSlider, createChoiceSlider } from '../../fields/sliderFieldBuilder';

import { getSizeLabel, getCountLabel, getCommaSepLabel } from '@mikezimm/npmfunctions/dist/Services/Math/basicOperations';

import { getSearchedFiles } from '../../ExSearch';

import { createItemsHeadingWithTypeIcons } from '../miniComps/components';

import { createItemDetail, getItemSearchString, getEventSearchString, getHighlightedText } from './SingleItem';

import { IItemSharingInfo, ISharingEvent, ISharedWithUser } from '../../Sharing/ISharingInterface';

const cellMaxStyle: React.CSSProperties = {
  whiteSpace: 'nowrap',
  overflow: 'hidden',
  maxWidth: '70%',
  height: '10px',
  textOverflow: 'ellipsis',
  cursor: 'pointer',
};

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

  private commonPath: string = this.getRelativePath === true && this.commonFolders.length > 0 ? this.commonRelativePath.replace( this.props.pickedList.LibraryUrl , '') + '/' : '';
  private commonParent: string = this.getRelativePath === true && this.commonFolders.length > 0 && this.commonRelativePath !== this.props.pickedList.LibraryUrl ? this.commonFolders[ this.commonFolders.length - 1 ] : '';

  private itemsHeading: any = createItemsHeadingWithTypeIcons( this.itemsAny, this.props.itemType, this.props.heading, this.props.icons );

  private sliderTitle = this.itemsLength < 400 ? 'Show Top items by size' : `Show up to 400 of ${ getCommaSepLabel(this.itemsLength) } items, use Search box to find more)`;
  private sliderMax = this.itemsLength < 400 ? this.itemsLength : 400;
  private sliderInc = this.itemsLength < 50 ? 1 : this.itemsLength < 100 ? 10 : 25;
  private siderMin = this.sliderInc > 1 ? this.sliderInc : 5;

  private searchMedia = this.props.dataOptions.useMediaTags !== true ? '' : ', MediaServiceAutoTags, MediaServiceKeyPoints, MediaServiceLocation, MediaServiceOCR';
  private searchNote = `Search will search Created Name and Date, filenames/types ${ this.searchMedia }`;

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
  this.props.items.map( item => {
    totalSize += item.size;
  });

  
  if ( this.props.itemType === 'Duplicates' ) {
    this.searchNote = `Search will search Created Name and Date, foldername/types`;

  } else if ( this.props.itemType === 'Shared' ) {
    this.searchNote = `Search will search date/time, sharedBy, sharedWith, filenames}`;
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

  public render(): React.ReactElement<IEsItemsProps> {

    console.log('EsItems.tsx1');
    // debugger;
    // const items : IItemDetail[] | IDuplicateFile []= this.itemsAny;
    let itemsTable = null;
    if ( this.props.itemType === 'Shared' ) {
      itemsTable = this.buildSharedEventsTable( this.sharedEvents , '', this.state.rankSlider, this.state.textSearch );

    } else if ( this.props.itemType === 'Items' ) {
      itemsTable = this.buildItemsTable( this.props.items , this.props.itemsAreDups, '', this.state.rankSlider, this.state.textSearch, 'size' );

    } else if ( this.props.itemType === 'Duplicates' ) {
        itemsTable = this.buildDupsTable( this.props.duplicateInfo.duplicates , this.props.itemsAreDups, '', this.state.rankSlider, this.state.textSearch, 'size' );

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

      } else if ( this.state.selectedItem ) { 

        panelContent = createItemDetail( this.state.selectedItem, this.props.itemsAreDups, this.props.pickedWeb.url, this.state.textSearch, this._onCloseItemDetail.bind( this ), this._onPreviewClick.bind( this ) );
    
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

      let panelStyle = this.showHeading !== true ? { marginTop: '1.4em'} : null;

      let styleCommonPathDisplay = this.commonPath === '' ? 'none' : null ;

      page = <div style= { panelStyle }>
        { this.showHeading !== true ? null : this.itemsHeading }
        <div className={ styles.inflexWrapCenter}>
          <div> { sliderTypeCount } </div>
          <div> { this.buildSearchBox( this.state.textSearch) } </div>
        </div>
        <div>
          <div>{ this.searchNote }</div>
          <div style={{ padding: '10px 0px 5px 0px', display: styleCommonPathDisplay }}>
            { `All items are below this folder: ${ this.props.pickedList.LibraryUrl.replace( this.props.pickedWeb.ServerRelativeUrl, '') }` }
            <span style={{ fontWeight: 600 }}>{ this.commonPath }</span></div>
        </div>
        { component }
      </div>;
    }

    return (
      <div className={ styles.exStorage } style={{ marginLeft: '25px'}}>
        { page }
        { userPanel }
      </div>
    );
  }

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
        { `Search all ${ getCommaSepLabel( this.props.items.length) } items [ ${ getSizeLabel( this.state.totalSize ) } ]` }
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

  private buildAllSharedEventsItems( items: IItemDetail[] ) {

    let sharedEvents: ISharingEvent[] = [];

    //Get all events
    items.map( item => {

      if ( item.itemSharingInfo && item.itemSharingInfo.sharedEvents ) {
        item.itemSharingInfo.sharedEvents.map( event => {
          event.ServerRedirectedEmbedUrl = item.ServerRedirectedEmbedUrl;
          event.parentFolder = item.parentFolder;
          sharedEvents.push( event );
        });
      }
    });

    //Sort by Shared Time
    sharedEvents = sortObjectArrayByChildNumberKey( sharedEvents, 'asc', 'TimeMS' );

    return sharedEvents;

  }

  private buildSharedEventsTable( sharedEvents: ISharingEvent[] , data: string, countToShow: number, textSearch: string, ): any {

    // let items : IItemDetail[] = itemsIn;
    let rows = [];
    let tableTitle = data;

    let priorEvent =  null;
    //Get event rows (if visible )
    sharedEvents.map( ( event, index ) => {
      if ( rows.length < countToShow ) {
        if ( this.isEventVisible( textSearch, event ) === true ) {
          rows.push( this.createSingleEventRow( index.toFixed(0), event , priorEvent, textSearch ) );
          priorEvent = event ;
        }
      }
    });

    return this.buildTableFromRows( rows, tableTitle, itemStyles.eventsTable );

  }

  private buildItemsTable( items: IItemDetail[] | IDuplicateFile[] , itemsAreDups: boolean , data: string, countToShow: number, textSearch: string, sortKey: 'size' ): any {

    let rows = [];
    let tableTitle = data;
    let itemsSorted: any[] = [];

    itemsSorted = sortObjectArrayByChildNumberKey( items, 'dec', sortKey );
    
    itemsSorted.map( ( item, index ) => {
      if ( rows.length < countToShow ) {
        if ( this.isVisibleItem( textSearch, item, itemsAreDups ) === true ) {
            rows.push( this.createSingleItemRow( index.toFixed(0), item ) );

        }
      }
    });

    return this.buildTableFromRows( rows, tableTitle, itemStyles.itemsTable );

  }

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

    itemsSorted = sortObjectArrayByChildNumberKey( items, 'dec', sortKey );
    
    itemsSorted.map( ( item, index ) => {
      if ( rows.length < countToShow ) {
        if ( this.isVisibleItem( textSearch, item, itemsAreDups ) === true ) {
            rows.push( this.createSingleDupRow( index.toFixed(0), item ) );
        }
      }
    });

    return this.buildTableFromRows( rows, tableTitle, itemStyles.itemsTable );

  }
  
  private isEventVisible ( textSearch: string, event: ISharingEvent ) {

    let visible = true;

    if ( textSearch.length > 0 ) {

      visible = false;
      let searchThis = getEventSearchString( event );
      if ( searchThis.toLowerCase().indexOf( textSearch.toLowerCase()) > -1 ) {
        visible = true;
      }
    }

    return visible;

  }

  private isVisibleItem ( textSearch: string, item: IItemDetail, itemsAreDups: boolean ) {

    let visible = true;

    if ( textSearch.length > 0 ) {

      visible = false;

      let searchThis = getItemSearchString( item, itemsAreDups );

      if ( searchThis.toLowerCase().indexOf( textSearch.toLowerCase()) > -1 ) {
        visible = true;

      } else if ( item.MediaServiceAutoTags && textSearch.toUpperCase() === 'MSAT' ) {
        visible = true;

      } else if ( item.MediaServiceKeyPoints && textSearch.toUpperCase() === 'MSKP' ) {
        visible = true;

      } else if ( item.MediaServiceLocation && textSearch.toUpperCase() === 'MSL' ) {
        visible = true;

      } else if ( item.MediaServiceOCR && textSearch.toUpperCase() === 'MSOCR' ) {
        visible = true;

      } else if ( item.isFolder === true && textSearch.toUpperCase() === 'ISFOLDER' ) {
        visible = true;

      } else if ( item.itemSharingInfo && textSearch.toUpperCase() === 'SHARED' ) {
        visible = true;

      } else if ( item.uniquePerms === true ) {
        if ( textSearch.toUpperCase() === 'UNIQUE' || textSearch.toUpperCase() === 'PERMS'  ) {
          visible = true;
        }
      } else if ( textSearch.toUpperCase() === 'USER<>' ) {
        if ( item.authorTitle !== item.editorTitle ) {
          visible = true;
        }
 
      } else if ( textSearch.toUpperCase() === 'USER=' ) {
        if ( item.authorTitle === item.editorTitle ) {
          visible = true;
        }

      } else {

      }

    }

    return visible;

  }

  
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

    cells.push( this.buildOpenItemCell( item, item.FileLeafRef, item.FileLeafRef ) );
  
    let cellRow = <tr> { cells } </tr>;

    return cellRow;
  
  }

  private createSingleItemRow( key: string, item: IItemDetail ) {

    let itemFolder : any = this.props.itemsAreFolders === true ? item : null;
    let folder: IFolderDetail = itemFolder;

    let created = new Date( item.created );
    let modified = new Date( item.modified );

    let cells : any[] = [];
    cells.push( <td style={{width: '50px'}} >{ key }</td> );

    let id = this.props.itemsAreDups === true ? item.parentFolder : item.FileLeafRef  ;
    let detailItemIcon = this.buildDetailIcon( item, id );

    let userStyle: any =  { width: '150px' } ;
    let userTitle = null;

    if ( item.authorTitle !== item.editorTitle ) { 
      userStyle.color = 'red';
      userStyle.fontWeight = 600;
      userTitle = `Edited by ${ item.editorTitle }`;
    }

    cells.push( detailItemIcon );

    if ( this.props.itemsAreFolders === true ) {
      cells.push( <td style={{width: '80px'}} >{ getCountLabel( folder.directCount, 0 ) }</td> );
      console.log('getting sizeLabel: ', folder );
      cells.push( <td style={{width: '80px'}} >{ getSizeLabel( folder.directSize ) }</td> );

    } else {
      cells.push( <td style={{width: '80px'}} >{ getSizeLabel( item.size ) }</td> );
    }

    cells.push( <td style={ userStyle } title={ userTitle }>{ item.authorTitle }</td> );
    let dateStyle : React.CSSProperties = {width: '160px'};
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

    // cells.push( <td style={cellMaxStyle}><a href={ item.FileRef } target={ '_blank' }>{ item.FileLeafRef }</a></td> );
    if ( this.props.itemsAreFolders === false ) {
      cells.push( this.buildFolderIcon( item ) );
      cells.push( <td style={{width: '60px'}} >{ item.versionlabel }</td> );
    }

    let cellText: any = item.FileLeafRef;
    //For duplicate files, this will show the relative path.
    //BUT to help when there are deep folders, it will show based on the common parent folder, not full folder url because it can be to long
    if ( this.props.itemsAreDups === true ) {
      cellText = <span><span style={{ fontWeight: 600 }}>{'../' + this.commonParent }</span><span>{ item.parentFolder.replace( this.commonRelativePath, '' ) } </span></span> ; //commonParent
    } 
    cells.push( this.buildOpenItemCell( item, item.id.toFixed(0) , cellText ) );
  
    let cellRow = <tr style={{ height: '27px' }}> { cells } </tr>;

    return cellRow;
  
  }

  private createSingleEventRow( key: string, event: ISharingEvent, priorEvent: ISharingEvent, highlight: string ) {

    let cells : any[] = [];
    cells.push( <td style={{width: '50px'}} >{ key }</td> );

    let dateStyle : React.CSSProperties = {width: '160px'};

    let eventTime: any = event.SharedTime.toLocaleString();
    let eventTimeTitle = event.SharedTime.toLocaleString();

    let sharedBy: any = event.sharedBy;
    let sharedWith: any = event.sharedWith;
    let FileLeafRef: any = event.FileLeafRef;
    let dateSearch = event.SharedTime.toLocaleDateString();

    let isSameEvent = priorEvent !== null && event.TimeMS === priorEvent.TimeMS && event.sharedBy === priorEvent.sharedBy && event.FileLeafRef === priorEvent.FileLeafRef ? true : false;

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

    cells.push( <td style={ dateStyle } title={ eventTimeTitle } onClick = { () => this._onCTRLClickSearch(dateSearch) }>{ eventTime }</td> );
    cells.push( <td style={ null } title={ event.sharedBy } onClick = { () => this._onCTRLClickSearch(event.sharedBy) } >{ sharedBy } </td> );
    cells.push( <td style={ null } title={ null } onClick = { () => this._onCTRLClickSearch(event.sharedWith) } >{ sharedWith } </td> );
    cells.push( <td style={ null } title={  event.FileLeafRef  } onClick = { () => this._onCTRLClickSearch(event.FileLeafRef) } >{ FileLeafRef }</td> );

    let cellText: any = event.FileLeafRef;
  
    let cellRow = <tr style={{ height: '27px' }}> { cells } </tr>;

    return cellRow;
  
  }

  private _onCTRLClickSearch( searchThis: string ) : void {
    //This is a quick "clear search" feature
    if ( searchThis === this.state.textSearch ) { searchThis = ''; }
    this._searchForItems ( searchThis );

  }

  private _onClickFolder( event ) {
    // console.log( event );
    console.log( event.currentTarget.id );
    let clickThisItem = parseInt(event.currentTarget.id);

    this.props.items.map( item => {
      if ( item.id === clickThisItem ) { 
        window.open( item.parentFolder, "_blank");
      }
    });
  }

  private _onClickItem( event ) {
    // console.log( event );
    console.log( event.currentTarget.id );

    if ( this.props.itemType === 'Items' ) {
      let clickThisItem = parseInt(event.currentTarget.id) ;
      let items: IItemDetail[] = this.props.items;
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
  
  private buildOpenItemCell ( item: IItemDetail | IDuplicateFile, itemId: string, text: any ) {
    let cell = <td style={cellMaxStyle} onClick={ this._onClickItem.bind(this)} 
    id={ itemId } 
    title={ `Item ID: ${ itemId }`}
  >
    { <Icon iconName= { item.iconName } style={ { fontSize: 'larger', color: item.iconColor, padding: '0px 15px 0px 0px', } }></Icon> }
    { text }</td>;


    return cell;
  }

  private buildFolderIcon ( item: IItemDetail ) {

    let iconCell = <td className = { itemStyles.folderIcons } 
      onClick={ this._onClickFolder.bind(this)} id={ item.id.toFixed(0) }
      title={ `Go to parent folder: ${ item.parentFolder }`}
      >
      {/* { <Icon iconName= {'FabricMovetoFolder'} style={{ padding: '4px 4px', fontSize: 'large' }}></Icon> } */}
      { fpsAppIcons.GoToFolder }
    </td>;
    return iconCell;

  }

  private buildDetailIcon ( item: IItemDetail, id: string ) {
    
    let MediaIcons: any[] = item.isMedia ? this.buildMediaIcons( item ) : [];

    let detailIcon = item.isMedia === true ? fpsAppIcons.ImageSearchRed : fpsAppIcons.DocumentSearch;

    if ( item.itemSharingInfo ) { detailIcon = fpsAppIcons.SharedItem; }
    else if ( item.uniquePerms === true ) { detailIcon = fpsAppIcons.UniquePerms; }

    let iconCell = <td className = { itemStyles.tableIcons } style={{width: '70px', cursor: 'pointer', position: 'relative' }} 
      onClick={ this._onClickItemDetail.bind(this)} id={ id }
      title={ `See all Item Details.` }
      >
      { detailIcon }
      <div style={{ display: 'inline-block', position: 'absolute', marginLeft: '3px' }}> { MediaIcons } </div>
    </td>;

    return iconCell;

  }

  private buildMediaIcons( item: any ) { //IItemDetail
    let MediaIcons = [];
    if ( item.isMedia ) {

      if ( item.MediaServiceOCR ) {
        MediaIcons.push(  <Icon iconName= { 'CircleShapeSolid' } style={{ top: '2px', left: '2px', fontSize: '6px', position: 'absolute', color: 'dimgray' }} title="MediaServiceOCR"></Icon> );
      }
      if ( item.MediaServiceAutoTags ) {
        MediaIcons.push(  <Icon iconName= { 'TagSolid' } style={{ top: '1px', left: '12px', fontSize: '9px', position: 'absolute', color: 'dimgray' }} title="MediaServiceAutoTags"></Icon> );
      }
      if ( item.MediaServiceKeyPoints ) {
        MediaIcons.push(  <Icon iconName= { 'Location' } style={{ top: '10px', left: '2px', fontSize: '5px', position: 'absolute', color: 'dimgray' }} title="MediaServiceKeyPoints"></Icon> );
      }
      if ( item.MediaServiceLocation ) {
        MediaIcons.push(  <Icon iconName= { 'POISolid' } style={{ top: '11px', left: '12px', fontSize: '8px', position: 'absolute', color: 'dimgray' }} title="MediaServiceLocation"></Icon> );
      }

    }

    return MediaIcons;

  }
  
  private _typeSlider(newValue: number){
    this.setState({
      rankSlider: newValue,
    });
  }

  private _onClickItemDetail( event ){
    console.log( event );
    console.log( event.currentTarget.id );
    let showThisType = event.currentTarget.id;
    let selectedItem = null;

    if ( this.props.itemType === 'Items' ) {
      this.props.items.map( item => {
        let checkThis = this.props.itemsAreDups === true ? item.parentFolder : item.FileLeafRef  ;
        if ( checkThis === showThisType ) { selectedItem = item ; }
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

  private _onDialogDismiss( event ) {
    this.setState({
      showItem: false,
      showPreview: false,
      selectedItem: null,
      showItems: [],
    });
  }

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
