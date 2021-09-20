import * as React from 'react';
import styles from '../../ExStorage.module.scss';
import { IEsItemsProps } from './IEsItemsProps';
import { IEsItemsState } from './IEsItemsState';
import { IExStorageState, IEXStorageList, IEXStorageBatch, IBatchData, IUserSummary, IFileType, IItemDetail } from '../../IExStorageState';
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
import { Icon  } from 'office-ui-fabric-react/lib/Icon';

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

import { createSingleItemRow } from './SingleItem';

export default class EsItems extends React.Component<IEsItemsProps, IEsItemsState> {

  private currentDate = new Date();
  private currentYear = this.currentDate.getFullYear();
  
  private sliderTitle = this.props.items.length < 400 ? 'Show Top items by size' : `Show up to 400 of ${this.props.items.length} items, use Search box to find more)`;
  private sliderMax = this.props.items.length < 400 ? this.props.items.length : 400;
  private sliderInc = this.props.items.length < 50 ? 1 : this.props.items.length < 100 ? 10 : 25;
  private siderMin = this.sliderInc > 1 ? this.sliderInc : 5;

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

  this.state = {

        isLoaded: true,
        isLoading: false,
        errorMessage: '',

        hasError: false,
      
        showPane: false,

        items: [],

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
    const itemsTable = this.buildItemsTable( this.props.items , '', this.state.rankSlider, this.state.textSearch, 'size' );

    let component = <div className={ styles.inflexWrapCenter}>
      { itemsTable }
    </div>;

    let sliderTypeCount = this.props.items.length < 5 ? null : 
      <div style={{margin: '0px 50px 20px 50px'}}> { createSlider( this.sliderTitle , this.state.rankSlider , this.siderMin, this.sliderMax, this.sliderInc , this._typeSlider.bind(this), this.state.isLoading, 350) }</div> ;

    let iconArray = this.props.icons.map( icon => {
      return ( <Icon iconName= { icon.iconName } title={ icon.iconTitle } style={ { fontSize: '24px', color: icon.iconColor, padding: '0px 0px 0px 15px', } }></Icon> );
    });

    let page = null;
    const emptyItemsElements = this.props.emptyItemsElements;

    if ( this.props.items.length === 0 && emptyItemsElements && emptyItemsElements.length > 0 ) {
      page = emptyItemsElements[Math.floor(Math.random()*emptyItemsElements.length)];  //https://stackoverflow.com/a/5915122

    } else {
      page = <div>
        <div className={styles.flexWrapStart}>
          <h3>{ this.props.items.length } Items found { this.props.heading }</h3> < div> { iconArray } </div>
        </div>
        <div className={ styles.inflexWrapCenter}>
          <div> { sliderTypeCount } </div>
          <div> { this.buildSearchBox() } </div>
        </div>
        { component }
      </div>;
    }

    return (
      <div className={ styles.exStorage } style={{ marginLeft: '25px'}}>
        { page }
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
        { `Search all ${ this.props.items.length } items` }
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

  private buildItemsTable( items: IItemDetail[] , data: string, countToShow: number, textSearch: string, sortKey: 'size' ): any {

    let rows = [];
    let tableTitle = data;
    const itemsSorted: IItemDetail[] = sortObjectArrayByNumberKey( items, 'dec', sortKey );

    itemsSorted.map( ( item, index ) => {
      if ( rows.length < countToShow ) {
        if ( textSearch.length > 0 ) {
          let createdDate = new Date( item.created );
          let searchThis = [item.FileLeafRef, item.authorTitle, item.editorTitle, createdDate.toLocaleDateString() ].join('|');
          if ( searchThis.toLowerCase().indexOf( textSearch.toLowerCase()) > -1 ) {
            rows.push( this.createSingleItemRow( index.toFixed(0), item ) );
          }
        } else {
          rows.push( this.createSingleItemRow( index.toFixed(0), item ) );
        }

      }
    });

    let table = <div style={{marginRight: '10px'}}>
      <h3 style={{ textAlign: 'center' }}> { tableTitle }</h3>
      {/* <table style={{padding: '0 20px'}}> */}
      <table style={{ tableLayout:"fixed", width:"80%" }} id="Select-b">
        { rows }
        {/* { itemsSorted.map( ( item, index ) => {
            return this.createSingleItemRow( item.id.toFixed(0), item );
          })
        } */}
      </table>
    </div>;
    return table;

  }

  private createSingleItemRow( key: string, item: IItemDetail ) {

    let created = new Date( item.created );

    let cellMaxStyle: React.CSSProperties = {
        whiteSpace: 'nowrap',
        overflow: 'hidden',
        maxWidth: '70%',
        height: '10px',
        textOverflow: 'ellipsis',
        cursor: 'pointer',
    };

    let cells : any[] = [];
    cells.push( <td style={{width: '50px'}} >{ key }</td> );
    cells.push( <td style={{width: '100px'}} >{ getSizeLabel( item.size ) }</td> );
    cells.push( <td style={{width: '150px'}} >{ item.authorTitle }</td> );
    cells.push( <td style={{width: '200px'}} >{ created.toLocaleString() }</td> );
    cells.push( <td style={{width: '50px', cursor: 'pointer' }} 
      onClick={ this._onClickFolder.bind(this)} id={ item.id.toFixed(0) }
      title={ `Go to parent folder: ${ item.parentFolder }`}
      >
      { <Icon iconName= {'FabricFolderSearch'} style={{ padding: '0px 4px', fontSize: 'large' }}></Icon> }
    </td> );  
    // cells.push( <td style={cellMaxStyle}><a href={ item.FileRef } target={ '_blank' }>{ item.FileLeafRef }</a></td> );


    const iconStyles: any = { root: {
      fontSize: 'larger',
      color: item.iconColor,
      padding: '0px 4px 0px 10px',
    }};

    cells.push( <td style={cellMaxStyle} onClick={ this._onClickItem.bind(this)} 
        id={ item.id.toFixed(0) } 
        title={ `Item ID: ${item.id}`}
      >
        { <Icon iconName= { item.iconName } style={ { fontSize: 'larger', color: item.iconColor, padding: '0px 15px 0px 0px', } }></Icon> }
        { item.FileLeafRef }</td> );
  
    let cellRow = <tr> { cells } </tr>;

    return cellRow;
  
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
    let clickThisItem = parseInt(event.currentTarget.id);

    this.props.items.map( item => {
      let openThisLink =  item.FileRef;
      if ( item.id === clickThisItem ) { window.open( openThisLink, "_blank"); }
    });
  }
  
  private _typeSlider(newValue: number){
    this.setState({
      rankSlider: newValue,
    });
  }

}
