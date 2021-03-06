/***
 *    d888888b .88b  d88. d8888b.  .d88b.  d8888b. d888888b       .d88b.  d88888b d88888b d888888b  .o88b. d888888b  .d8b.  db      
 *      `88'   88'YbdP`88 88  `8D .8P  Y8. 88  `8D `~~88~~'      .8P  Y8. 88'     88'       `88'   d8P  Y8   `88'   d8' `8b 88      
 *       88    88  88  88 88oodD' 88    88 88oobY'    88         88    88 88ooo   88ooo      88    8P         88    88ooo88 88      
 *       88    88  88  88 88~~~   88    88 88`8b      88         88    88 88~~~   88~~~      88    8b         88    88~~~88 88      
 *      .88.   88  88  88 88      `8b  d8' 88 `88.    88         `8b  d8' 88      88        .88.   Y8b  d8   .88.   88   88 88booo. 
 *    Y888888P YP  YP  YP 88       `Y88P'  88   YD    YP          `Y88P'  YP      YP      Y888888P  `Y88P' Y888888P YP   YP Y88888P 
 *                                                                                                                                  
 *                                                                                                                                  
 */

import * as React from 'react';

import { escape } from '@microsoft/sp-lodash-subset';

import { Spinner, SpinnerSize, SpinnerLabelPosition } from 'office-ui-fabric-react/lib/Spinner';
import { Stack, IStackStyles, IStackTokens } from 'office-ui-fabric-react/lib/Stack';

import { IconButton, IIconProps, IContextualMenuProps, Link } from 'office-ui-fabric-react';

import { Panel, IPanelProps, IPanelStyleProps, IPanelStyles, PanelType } from 'office-ui-fabric-react/lib/Panel';

import {
  MessageBar,
  MessageBarType,
  SearchBox,
  Icon,
  Label,
  Pivot,
  PivotItem,
  PivotLinkFormat,
  PivotLinkSize,
  Dropdown,
  IDropdownOption,
} from "office-ui-fabric-react";

import { IGrid } from 'office-ui-fabric-react';

/***
 *    d888888b .88b  d88. d8888b.  .d88b.  d8888b. d888888b      d8b   db d8888b. .88b  d88.      d88888b db    db d8b   db  .o88b. d888888b d888888b  .d88b.  d8b   db .d8888. 
 *      `88'   88'YbdP`88 88  `8D .8P  Y8. 88  `8D `~~88~~'      888o  88 88  `8D 88'YbdP`88      88'     88    88 888o  88 d8P  Y8 `~~88~~'   `88'   .8P  Y8. 888o  88 88'  YP 
 *       88    88  88  88 88oodD' 88    88 88oobY'    88         88V8o 88 88oodD' 88  88  88      88ooo   88    88 88V8o 88 8P         88       88    88    88 88V8o 88 `8bo.   
 *       88    88  88  88 88~~~   88    88 88`8b      88         88 V8o88 88~~~   88  88  88      88~~~   88    88 88 V8o88 8b         88       88    88    88 88 V8o88   `Y8b. 
 *      .88.   88  88  88 88      `8b  d8' 88 `88.    88         88  V888 88      88  88  88      88      88b  d88 88  V888 Y8b  d8    88      .88.   `8b  d8' 88  V888 db   8D 
 *    Y888888P YP  YP  YP 88       `Y88P'  88   YD    YP         VP   V8P 88      YP  YP  YP      YP      ~Y8888P' VP   V8P  `Y88P'    YP    Y888888P  `Y88P'  VP   V8P `8888Y' 
 *                                                                                                                                                                              
 *                                                                                                                                                                              
 */

import {  getOffSetDayOfWeek, getYearWeekLabel} from '@mikezimm/npmfunctions/dist/Services/Time/weeks';

import { getYearMonthLabel } from '@mikezimm/npmfunctions/dist/Services/Time/getLabels';

import { monthStr3, } from '@mikezimm/npmfunctions/dist/Services/Time/monthLabels';
import { weekday3 } from '@mikezimm/npmfunctions/dist/Services/Time/dayLabels';

import { getTimeDelta } from '@mikezimm/npmfunctions/dist/Services/Time/deltas';

import { sortObjectArrayByStringKey } from '@mikezimm/npmfunctions/dist/Services/Arrays/sorting';

import { getExpandColumns, getSelectColumns, IPerformanceSettings, createFetchList, IZBasicList, } from '@mikezimm/npmfunctions/dist/Lists/getFunctions';

import { getSizeLabel, getCountLabel, getCommaSepLabel } from '@mikezimm/npmfunctions/dist/Services/Math/basicOperations';

/***
 *    d888888b .88b  d88. d8888b.  .d88b.  d8888b. d888888b      .d8888. d88888b d8888b. db    db d888888b  .o88b. d88888b .d8888. 
 *      `88'   88'YbdP`88 88  `8D .8P  Y8. 88  `8D `~~88~~'      88'  YP 88'     88  `8D 88    88   `88'   d8P  Y8 88'     88'  YP 
 *       88    88  88  88 88oodD' 88    88 88oobY'    88         `8bo.   88ooooo 88oobY' Y8    8P    88    8P      88ooooo `8bo.   
 *       88    88  88  88 88~~~   88    88 88`8b      88           `Y8b. 88~~~~~ 88`8b   `8b  d8'    88    8b      88~~~~~   `Y8b. 
 *      .88.   88  88  88 88      `8b  d8' 88 `88.    88         db   8D 88.     88 `88.  `8bd8'    .88.   Y8b  d8 88.     db   8D 
 *    Y888888P YP  YP  YP 88       `Y88P'  88   YD    YP         `8888Y' Y88888P 88   YD    YP    Y888888P  `Y88P' Y88888P `8888Y' 
 *                                                                                                                                 
 *                                                                                                                                 
 */


 /***
 *    d888888b .88b  d88. d8888b.  .d88b.  d8888b. d888888b      db   db d88888b db      d8888b. d88888b d8888b. .d8888. 
 *      `88'   88'YbdP`88 88  `8D .8P  Y8. 88  `8D `~~88~~'      88   88 88'     88      88  `8D 88'     88  `8D 88'  YP 
 *       88    88  88  88 88oodD' 88    88 88oobY'    88         88ooo88 88ooooo 88      88oodD' 88ooooo 88oobY' `8bo.   
 *       88    88  88  88 88~~~   88    88 88`8b      88         88~~~88 88~~~~~ 88      88~~~   88~~~~~ 88`8b     `Y8b. 
 *      .88.   88  88  88 88      `8b  d8' 88 `88.    88         88   88 88.     88booo. 88      88.     88 `88. db   8D 
 *    Y888888P YP  YP  YP 88       `Y88P'  88   YD    YP         YP   YP Y88888P Y88888P 88      Y88888P 88   YD `8888Y' 
 *                                                                                                                       
 *                                                                                                                       
 */


import { createSlider, createChoiceSlider } from '../../fields/sliderFieldBuilder';

import {updateAllItems, IGridList } from './GetListData';

 /***
 *    d888888b .88b  d88. d8888b.  .d88b.  d8888b. d888888b       .o88b.  .d88b.  .88b  d88. d8888b.  .d88b.  d8b   db d88888b d8b   db d888888b 
 *      `88'   88'YbdP`88 88  `8D .8P  Y8. 88  `8D `~~88~~'      d8P  Y8 .8P  Y8. 88'YbdP`88 88  `8D .8P  Y8. 888o  88 88'     888o  88 `~~88~~' 
 *       88    88  88  88 88oodD' 88    88 88oobY'    88         8P      88    88 88  88  88 88oodD' 88    88 88V8o 88 88ooooo 88V8o 88    88    
 *       88    88  88  88 88~~~   88    88 88`8b      88         8b      88    88 88  88  88 88~~~   88    88 88 V8o88 88~~~~~ 88 V8o88    88    
 *      .88.   88  88  88 88      `8b  d8' 88 `88.    88         Y8b  d8 `8b  d8' 88  88  88 88      `8b  d8' 88  V888 88.     88  V888    88    
 *    Y888888P YP  YP  YP 88       `Y88P'  88   YD    YP          `Y88P'  `Y88P'  YP  YP  YP 88       `Y88P'  VP   V8P Y88888P VP   V8P    YP    
 *                                                                                                                                               
 *                                                                                                                                               
 */

import styles from './Gridcharts.module.scss';
import { IGridchartsProps, IValueOperator,  } from './IGridchartsProps';
import { IGridchartsState, IGridchartsData, IGridchartsDataPoint, IGridItemInfo, ITimeScale, ISizeOrCount } from './IGridchartsState';

import EsItems from '../items/EsItems';

/***
 *    d88888b db    db d8888b.  .d88b.  d8888b. d888888b      d888888b d8b   db d888888b d88888b d8888b. d88888b  .d8b.   .o88b. d88888b .d8888. 
 *    88'     `8b  d8' 88  `8D .8P  Y8. 88  `8D `~~88~~'        `88'   888o  88 `~~88~~' 88'     88  `8D 88'     d8' `8b d8P  Y8 88'     88'  YP 
 *    88ooooo  `8bd8'  88oodD' 88    88 88oobY'    88            88    88V8o 88    88    88ooooo 88oobY' 88ooo   88ooo88 8P      88ooooo `8bo.   
 *    88~~~~~  .dPYb.  88~~~   88    88 88`8b      88            88    88 V8o88    88    88~~~~~ 88`8b   88~~~   88~~~88 8b      88~~~~~   `Y8b. 
 *    88.     .8P  Y8. 88      `8b  d8' 88 `88.    88           .88.   88  V888    88    88.     88 `88. 88      88   88 Y8b  d8 88.     db   8D 
 *    Y88888P YP    YP 88       `Y88P'  88   YD    YP         Y888888P VP   V8P    YP    Y88888P 88   YD YP      YP   YP  `Y88P' Y88888P `8888Y' 
 *                                                                                                                                               
 *                                                                                                                                               
 */


/**
 * Based upon example from
 * https://codepen.io/ire/pen/Legmwo
 */

export default class Gridcharts extends React.Component<IGridchartsProps, IGridchartsState> {

    private valueOperatorOptions :IDropdownOption[] = this.props.columns.valueOperators.map( operator => {
      return { key: operator, text: operator } ;
    });

    //https://stackoverflow.com/a/4413721
    private addDays (tempDate, days) {
      var date = new Date(tempDate.valueOf());
      date.setDate(date.getDate() + days);
      return date;
    }

    //https://stackoverflow.com/a/4413721
    private getDates(startDate, stopDate) {
      var dateArray = new Array();
      var currentDate = startDate;
      while (currentDate <= stopDate) {
          let tempDate = new Date (currentDate);
          dateArray.push(tempDate);
          currentDate = this.addDays( tempDate , 1);
      }
      return dateArray;
    }

    private createSampleGridData() {

      let arrDates: any[] = [];
      let startDate = new Date();
      let endDate = new Date();
      endDate.setDate(endDate.getDate() + 365 - 2 );

      arrDates = this.getDates( startDate, endDate);
      let allDataPoints : IGridchartsDataPoint[] = [];

      for (var i = 1; i < 365; i++) {

        let data : IGridchartsDataPoint = {
          date: null,
          dateNo: null,
          dayNo: null,
          week: null,
          month: null,
          year: null,
          yearMonth: null,
          yearWeek: null,

          yearIndex: null,
          yearMonthIndex: null,
          yearWeekIndex: null,

          label: null,
          dateString: '',
          dataLevel: Math.floor(Math.random() * 3),
          value: Math.floor(Math.random() * 20),
          values: [],
          valuesString: [],
          Count: null,
          Sum: null,
          Avg: null,
          Min: null,
          Max: null,
          items: [],
        };

        let thisDate : Date = arrDates[ i- 1];
        data.label = thisDate.toLocaleDateString();
        data.date = thisDate;
        allDataPoints.push( data ); 

      }
      return allDataPoints;
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


    public constructor(props:IGridchartsProps){
        super(props);
          /**
           * This is copied later in code when you have to call the data in case something changed.
           */  //createGridList(webURL, parentListURL, name, isLibrary, performance, pageContext, title: string = null)

          /*
          dateColumn: string;
          valueColumn: string;
          valueType: string;
          valueOperator: string;
        */
        let allColumns : string[] = [];
        let dropDownColumns: string[] = this.props.columns.dropDownColumns;
        let searchColumns : string[] = this.props.columns.searchColumns;
        let metaColumns : string[] = this.props.columns.metaColumns;
        let expandDates : string[] = [this.props.columns.dateColumn];
        let selectedDropdowns: string[] = [];

        
        allColumns.push( this.props.columns.dateColumn );
        allColumns.push( this.props.columns.valueColumn );

        searchColumns.map( c => { allColumns.push( c ) ; });
        metaColumns.map( c => { allColumns.push( c ) ; });

        let dropDownSort : string[] = dropDownColumns.map( c => { let c1 = c.replace('>','') ; if ( c1.indexOf('-') === 0 ) { return 'dec' ; } else if ( c1.indexOf('+') === 0 ) { return 'asc' ; } else { return ''; } });

        dropDownColumns.map( c => { let c1 = c.replace('>','').replace('+','').replace('-','') ; searchColumns.push( c1 ) ; metaColumns.push( c1 ) ; allColumns.push( c1 ); selectedDropdowns.push('') ; });

        let basicList : IZBasicList = createFetchList( this.props.parentListWeb, null, this.props.parentListTitle, null, null, this.props.performance, this.props.pageContext, allColumns, searchColumns, metaColumns, expandDates );
        //Have to do this to add dropDownColumns and dropDownSort to IZBasicList
        let tempList : any = basicList;
        tempList.dropDownColumns = dropDownColumns;
        tempList.dropDownSort = dropDownSort;
        let fetchList : IGridList = tempList;

        console.log('fetchList Contructor:', fetchList );
        /**
         * Add this at this point to be able to search on specific odata types
         * fetchList.odataSearch = ['odata.type'];
        */

        let errMessage = null;

        let allDataPoints : IGridchartsDataPoint[] = this.createSampleGridData();

        //console.log('gridData', allDataPoints );

        const s1 = allDataPoints[0].date.getMonth();
        const s2 = s1 + 12;

        const monthLables = monthStr3["en-us"].concat( ... monthStr3["en-us"] ).slice(s1,s2) ;
        const monthScales = [ 4,4,4,5,4,4,5,4,4,5,4,5   ,   4,4,4,5,4,4,5,4,4,5,4,5 ].slice(s1,s2) ;

        let allDateArray = [];

        
        let sizeOrCount : ISizeOrCount = this.props.columns.valueOperators[0] !== 'Count' || this.props.columns.valueColumn.toLowerCase().indexOf('size') > -1 ? 'size' : 'count';

        let gridData: IGridchartsData = {

          startDate: null,
          endDate: null,
          gridEnd: null,
          gridStart: null,

          allDataPoints: allDataPoints,
          allDateArray: allDateArray,
          allDateStringArray: [],

          allYearsStringArray: [],
          allMonthsStringArray: [],
          allWeekNosStringArray: [],

          allWeeks: 0,

          visibleDataPoints: [],
          visibleDateArray: [],
          visibleDateStringArray: [],
          visibleWeeks: 0,
          
          total: null,
          count: 0,
          leadingBlanks: 0,

          maxValue: null,
          minValue: null,

          totalLabel: '',
          sizeOrCount: sizeOrCount,

          dataLevelLabels: [],

        };

        this.state = { 

          //Size courtesy of https://www.netwoven.com/2018/11/13/resizing-of-spfx-react-web-parts-in-different-scenarios/
          WebpartHeight: this.props.WebpartElement ? this.props.WebpartElement.getBoundingClientRect().height : null,
          WebpartWidth:  this.props.WebpartElement ? this.props.WebpartElement.getBoundingClientRect().width - 50 : null,

          monthLables: monthLables,
          monthScales: monthScales,

          sliderValueWeek: 0,

          sliderValueYear: 0,
          sliderValueMonth: 0,
          sliderValueWeekNo: 0,

          timeSliderScale: [ 'Weeks', 'Years', 'Months', 'WeekNo'],
          currentTimeScale: 'Weeks',

          choiceSliderValue: 0,
          breadCrumb: [],
          choiceSliderDropdown: null,
          showChoiceSlider: false,

          dropdownColumnIndex: null,

          valueOperator: this.props.columns.valueOperators[0],

          selectedYear: null,
          selectedUser: null,
          selectedDropdowns: selectedDropdowns,
          dropDownItems: [],

          gridData: gridData,

          fetchList: fetchList,

          bannerMessage: null,
          showTips: false,

          allowOtherSites: this.props.allowOtherSites === true ? true : false,
          allLoaded: false,

          allItems: [],
          searchedItems: [],
          stats: [],
          first20searchedItems: [],
          searchCount: 0,

          meta: [],

          webURL: this.props.parentListWeb,

          searchMeta: null, // [pivCats.all.title],
          searchText: '',

          errMessage: errMessage,
          
          pivotCats: [],

          lastStateChange: 'Loading',
          stateChanges: [],

          showItems: -1,
//          style: this.props.style ? this.props.style : 'commandBar',

        };  

    }


    public componentDidMount() {

      console.log('fetchList componentDidMount:', this.state.fetchList );
      updateAllItems( this.props.items, this.state.fetchList, this.addTheseItemsToState.bind(this), null, null );
      
    }


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

      let reloadData : any = false;
      let refreshMe : any = false;

      let reloadOnThese = [
        'parentListTitle', 'parentListName', 'parentListWeb', '', '',
        'dateColumn', 'valueColumn', 'valueType', 'valueOperator','dropDownColumns',
      ];

      let refreshOnThese = [
        'setSize','setTab','otherTab','setTab','otherTab','setTab','otherTab','setTab','otherTab', '',
        'pivotSize', 'pivotFormat', 'pivotOptions', 'pivotTab', 'advancedPivotStyles', 'gridStyles',

      ];

      reloadOnThese.map( key => {
        if ( prevProps[key] !== this.props[key] ) { reloadData = true; }
      });

      if (reloadData === false) {
        refreshOnThese.map( key => {
          if ( prevProps[key] !== this.props[key] ) { refreshMe = true; }
        });
      }

      if ( this.props.items.length !== prevProps.items.length ) { reloadData = true ; }

      if ( this.props.refreshId !== prevProps.refreshId ) { reloadData = true; }

      if (reloadData === true) {
        //Need to first update fetchList and pass it on.

        let allColumns : string[] = [];
        let dropDownColumns: string[] = this.props.columns.dropDownColumns;
        let searchColumns : string[] = this.props.columns.searchColumns;
        let metaColumns : string[] = this.props.columns.metaColumns;
        let expandDates : string[] = [this.props.columns.dateColumn];
        
        allColumns.push( this.props.columns.dateColumn );
        allColumns.push( this.props.columns.valueColumn );

        searchColumns.map( c => { allColumns.push( c ) ; });
        metaColumns.map( c => { allColumns.push( c ) ; });

        let dropDownSort : string[] = dropDownColumns.map( c => { let c1 = c.replace('>','') ; if ( c1.indexOf('-') === 0 ) { return 'dec' ; } else if ( c1.indexOf('+') === 0 ) { return 'asc' ; } else { return ''; } });

        dropDownColumns.map( c => { let c1 = c.replace('>','').replace('+','').replace('-','') ; searchColumns.push( c1 ) ; metaColumns.push( c1 ) ; allColumns.push( c1 ); });

        let basicList : IZBasicList = createFetchList( this.props.parentListWeb, null, this.props.parentListTitle, null, null, this.props.performance, this.props.pageContext, allColumns, searchColumns, metaColumns, expandDates );
        //Have to do this to add dropDownColumns and dropDownSort to IZBasicList
        let tempList : any = basicList;
        tempList.dropDownColumns = dropDownColumns;
        tempList.dropDownSort = dropDownSort;
        let fetchList : IGridList = tempList;

        console.log('fetchList componentDidUpdate:', fetchList );
        updateAllItems( this.props.items, fetchList, this.addTheseItemsToState.bind(this), null, null );
        
      } else if ( refreshMe === true ) {  this.setState({ }) ; }


      //if (this.props.defaultProjectPicker !== prevProps.defaultProjectPicker) {  rebuildTiles = true ; }

      //if (rebuildTiles === true) {
        //this._updateStateOnPropsChange({});
      //}
    }



/***
 *         d8888b. d88888b d8b   db d8888b. d88888b d8888b. 
 *         88  `8D 88'     888o  88 88  `8D 88'     88  `8D 
 *         88oobY' 88ooooo 88V8o 88 88   88 88ooooo 88oobY' 
 *         88`8b   88~~~~~ 88 V8o88 88   88 88~~~~~ 88`8b   
 *         88 `88. 88.     88  V888 88  .8D 88.     88 `88. 
 *         88   YD Y88888P VP   V8P Y8888D' Y88888P 88   YD 
 *                                                          
 *                                                          
 */

  public render(): React.ReactElement<IGridchartsProps> {

    const wrapStackTokens: IStackTokens = { childrenGap: 30 };

    let gridElement = null;
    let searchStack = null;
    let showChoiceSlider = this.state.showChoiceSlider;
    let sliderValueWeek = this.state.sliderValueWeek;
    let sliderValueYear = this.state.sliderValueYear;
    let sliderValueMonth = this.state.sliderValueMonth;
    let sliderValueWeekNo = this.state.sliderValueWeekNo;
    let currentTimeScale : ITimeScale = this.state.currentTimeScale;

    let yearSliderMax = ( this.state.gridData.allYearsStringArray.length -365 );
    let weekNoSliderMax = ( this.state.gridData.allWeekNosStringArray.length -365 );
    let monthSliderMax = ( this.state.gridData.allMonthsStringArray.length -365 );

    let monthLabels = [];
    let lastMonth = null;
    let yearLabels = [];
    let lastYear = null;
    //transparent,#ebedf0,#c6e48b,#7bc96f,#196127   li, -1, 1, 2, 3

    let cellColor = this.props.gridStyles.cellColor;


    let dataLevelli = { backgroundColor: 'transparent' };
    let dataLevelMinus1Style = { backgroundColor: this.props.gridStyles.emptyColor, opacity: 1, };
    let dataLevel0Style = { backgroundColor: 'transparent', };
    let dataLevel1Style = { backgroundColor: this.props.gridStyles.squareColor, opacity: .33, };
    let dataLevel2Style = { backgroundColor: this.props.gridStyles.squareColor, opacity: .66, };   
    let dataLevel3Style = { backgroundColor: this.props.gridStyles.squareColor, opacity: 1, };

    if ( this.props.gridStyles.cellColor === 'green' ) {
      //transparent,#ebedf0,#c6e48b,#7bc96f,#196127
      dataLevel0Style = { backgroundColor: '#ebedf0' };
      dataLevelMinus1Style = { backgroundColor: 'transparent', opacity: 1, };
      dataLevel1Style = { backgroundColor: '#c6e48b', opacity: 1, };
      dataLevel2Style = { backgroundColor: '#7bc96f', opacity: 1, };   
      dataLevel3Style = { backgroundColor: '#196127', opacity: 1, };

    } else if ( this.props.gridStyles.cellColor === 'custom' && this.props.gridStyles.squareCustom.length > 0 ) {
      let squareCustom = this.props.gridStyles.squareCustom.split(',');
      dataLevel0Style = { backgroundColor: squareCustom[0] };
      dataLevelMinus1Style = { backgroundColor: squareCustom[1], opacity: 1, };
      dataLevel1Style = { backgroundColor: squareCustom[2], opacity: 1, };
      dataLevel2Style = { backgroundColor: squareCustom[3], opacity: 1, };   
      dataLevel3Style = { backgroundColor: squareCustom[4], opacity: 1, };

    }

    let sliderTransform = null;
    let weekSliderMax = ( this.state.gridData.allDateArray.length -365 ) / 7 + 1;

    //Add extra 'weeks' or spaces for each month's gaps
    weekSliderMax += this.state.gridData.allMonthsStringArray.length * parseInt( this.props.gridStyles.monthGap );

    if ( weekSliderMax < 2 ) { weekSliderMax = 2 ; }

    const squares : any[] = [];

    if ( this.state.allLoaded === true ) {

      console.log('gridData Render:', this.state.gridData );
      /**
       * These loops add leading squares and must be before pushing actual data
       */
      if ( this.props.scaleMethod === 'slider') {
        //Do nothing special at this time
        sliderTransform = this.props.scaleMethod === 'slider' ? "translate3d(" + ( -sliderValueWeek ) + "vw, 0, 0)" : null;

      } else if ( this.props.scaleMethod === 'blink' && sliderValueWeek < 0 ) {
          for (let i = 1; i < sliderValueWeek * 7; i++) { //This just tests sliding grid animation
            squares.push(<li data-level={ -1 } style={ dataLevelMinus1Style } ></li>);
          }
          sliderTransform = '';
      }

      if ( this.state.gridData.leadingBlanks > 0 ) {
        for (let lb = 1; lb < this.state.gridData.leadingBlanks; lb++) { //this works for regular leading blanks
            squares.push(<li data-level={ -1 } style={ dataLevelMinus1Style }></li>);
          }
      }

      /**
       * This loop adds all the real squares to the mix
       */
      let pushSpacer = true;

      this.state.gridData.allDataPoints.map( ( d, index ) => {
        if ( this.props.scaleMethod === 'blink' && sliderValueWeek > 0 &&
            index < sliderValueWeek * 7 ) {
          //Skip drawing these squares (this week is to left of grid )
        } else if ( squares.length < 370 ) { //Only push 1 year's worth of items

          //This will add 7 days of white spaces between months
          let fillerDays = this.props.gridStyles.monthGap === "2" ? 14 : this.props.gridStyles.monthGap === "1" ? 7 : 0 ;
          if ( fillerDays > 0 && index !== 0 && d.dateNo === 1 ) {
            for (let day = 0; day < fillerDays; day++) { //this works for regular leading blanks
              squares.push(<li data-level={ -1 } style={ dataLevelMinus1Style }></li>);
            }
          }

          if ( d.dayNo === 0 ) { //This is a sunday, update MonthLabels
            if ( d.month !== lastMonth ) {
              if ( lastMonth !== null ) { //Add spacer weeks but Skip this on the first month
                if ( fillerDays >= 7 ) { monthLabels.push( null ); yearLabels.push( null );}
                if ( fillerDays >= 14 ) { monthLabels.push( null ); yearLabels.push( null ); }
              }
              lastMonth = d.month;
              monthLabels.push( monthStr3["en-us"][ lastMonth ] );

            } else {
              monthLabels.push( null );
            }
          } else if ( pushSpacer === true ) { //Add spacer if first day of range is not Sunday
            monthLabels.push( null );
            yearLabels.push( null );
            pushSpacer = false;
          }

          if ( d.dayNo === 0 ) { //This is a sunday, update MonthLabels
            if ( d.year !== lastYear ) {
              lastYear = d.year;
              if ( lastYear !== null ) { //Add spacer weeks but Skip this on the first month
                //if ( fillerDays >= 7 ) { yearLabels.push( null ); }
                //if ( fillerDays >= 14 ) { yearLabels.push( null ); }
              }
              yearLabels.push( d.year );

            } else {
              yearLabels.push( null );
            }
          }


          let thisStyle = null;
          if ( d.dataLevel === 0 ) { thisStyle = dataLevel0Style ; }
          else if ( d.dataLevel === 1 ) { thisStyle = dataLevel1Style ; }
          else if ( d.dataLevel === 2 ) { thisStyle = dataLevel2Style ; }
          else if ( d.dataLevel === 3 ) { thisStyle = dataLevel3Style ; }
          else { thisStyle = dataLevel3Style ; }

          if ( d.items.length > 0 ) {
            //Add onClick to open panel
            squares.push( <li className = { styles.pointer } style={ thisStyle } title={ d.label + ' : level ' + d.dataLevel } data-level={ d.dataLevel }
              onClick= {(ev) => {
                this._showTheseItems(d.label, index, ev);
              }}
            ></li> ) ;
          } else {
            squares.push( <li style={ thisStyle } title={ d.label + ' : ' + d.dataLevel } data-level={ d.dataLevel }></li> ) ;
          }


        }
      });

      /**
       * Adding overflow hidden on Squares limits visible squares to the width of the element.
       * BUT the entire year slides and is not trimmed by parent element size location... so the 1 year can slide over dates and off the screen.
       * Need to have something else mask it when it goes out of the visible area.
       * That would also mean having it not transparent so you have to fix the background color which may not match another color.
      */
      let backGroundColor = this.props.gridStyles.squareCustom.length > 0 ? this.props.gridStyles.squareCustom.split(',')[0] : this.props.gridStyles.backGroundColor;

      gridElement = <ul className={styles.squares} style={{ backgroundColor: backGroundColor, listStyleType: 'none', transform: sliderTransform, transition: 'transform .3s cubic-bezier(0, .52, 0, 1)' }}>
                        { squares }
                    </ul>;

      let searchElements = [];
      let choiceSlider = null;
      /**
       * Add Dropdown search
       */
        if ( this.state.dropDownItems.length > 0 ) {

          let choiceSliderDropdown = this.state.choiceSliderDropdown;
          if ( showChoiceSlider === true && choiceSliderDropdown !== null ) {
            let choiceSliderValue = this.state.choiceSliderValue;
            let choiceMax = this.state.dropDownItems[choiceSliderDropdown].length -1 ;
  
            if ( choiceSliderValue !== null ) {
              console.log('choiceSliderValue, this.state.dropDownItems:', choiceSliderValue, this.state.dropDownItems);
              console.log('choiceSliderDropdown, this.state.dropDownItems[choiceSliderDropdown]:', choiceSliderDropdown, this.state.dropDownItems[choiceSliderDropdown]);
              let theChoice = choiceSliderValue > -1 ? `${ this.state.dropDownItems[choiceSliderDropdown][choiceSliderValue].text } ` : 'TBD' ;
  
              choiceSlider = this.state.dropDownItems.length === 0 ? null : 
                <div><div style={{position: 'absolute', paddingTop: '10px', paddingLeft: '30px'}}>{ /* theChoice */  }</div>
                  <Stack horizontal horizontalAlign='center' >
                    <div style={{ width: '30%', paddingLeft: '50px', paddingRight: '50px', paddingTop: '10px' }}>
                      { createChoiceSlider('Slide to adjust choice', theChoice , choiceMax, 1 , this._updateChoiceSlider.bind(this)) }
                    </div>
                  </Stack></div>;
              
            }
          }

          searchElements = this.state.dropDownItems.map( ( dropDownChoices, index ) => {

              let dropDownSort = this.state.fetchList.dropDownSort[ index ];
              let dropDownChoicesSorted = dropDownSort === '' ? dropDownChoices : sortObjectArrayByStringKey( dropDownChoices, dropDownSort, 'text' );
              let DDLabel = this.state.fetchList.dropDownColumns[ index ].replace('>','').replace('+','').replace('-','');
              return <Dropdown
                  placeholder={ `Select a ${ DDLabel }` }
                  label={ DDLabel }
                  options={dropDownChoicesSorted}
                  selectedKey={ this.state.selectedDropdowns [index ] === '' ? null : this.state.selectedDropdowns [ index ] }
                  onChange={(ev: any, value: IDropdownOption) => {
                    this.searchForItems(value.key.toString(), index, DDLabel, ev);
                  }}
                  styles={{ dropdown: { width: 200 } }}
              />;
          });
        } 

        /**
         * Add Text search box
         */
        if ( this.props.enableSearch === true ) {

          let searchBox = <div>
          <div style={{ paddingTop: '20px' }}></div>
          <SearchBox className={ styles.searchBox }
              placeholder= { 'Search items' }
              iconProps={{ iconName : 'Search'}}
              onSearch={ this.textSearch.bind(this) }
              value={this.state.searchText}
              onChange={ this.textSearch.bind(this) }
               />
          </div>;

          searchElements.push( searchBox ) ;

        }

        let valueOpDropdown = <Dropdown
            placeholder={ `Select an operation` }
            label={ 'Operation' }
            options={ this.valueOperatorOptions }
            selectedKey={ this.state.valueOperator }
            onChange={(ev: any, value: IDropdownOption) => {
              this.changeOperation(value.key.toString(), ev);
            }}
            styles={{ dropdown: { width: 200 } }}
        />;

        searchElements.push( valueOpDropdown ) ;

        searchStack = <div style={{marginLeft: '38px'}}>
                <Stack horizontal horizontalAlign="start" verticalAlign="end" wrap tokens={wrapStackTokens}>
                  { searchElements }
                </Stack>
                <div> { choiceSlider } </div>
            </div>;

    } else {

      gridElement = <div style={{ position: 'absolute', top: '50%', left: '42%' }}>
          <Spinner 
            size={SpinnerSize.large}
            label={ 'Loading data' }
            labelPosition='left'
          ></Spinner>
        </div> ;
    }

    let metrics : any = <div className={ styles.metrics }>TBD</div>;
    if ( this.state.gridData.count > 0 ) {

      let line1 = `${ getCommaSepLabel( this.state.gridData.count) } items`;
      let line2 = `${ this.state.valueOperator} of ${ this.props.columns.valueColumn } = ${ this.state.gridData.totalLabel }`;
      metrics = <div className={ styles.metrics }>
          <div>{line1}</div>
          <div>{line2}</div>
      </div>;
    } 

    let timeSlider = null;
    if ( this.props.scaleMethod === 'slider' ||  this.props.scaleMethod === 'blink' ) {

      //let yearSliderMax = ( this.state.gridData.allYearsStringArray.length -365 );
      //let weekNoSliderMax = ( this.state.gridData.allWeekNosStringArray.length -365 );
      //let monthSliderMax = ( this.state.gridData.allMonthsStringArray.length -365 );

      let activeSlider = null;
      if ( currentTimeScale === 'Weeks' ) {

        activeSlider = createSlider('Slide to adjust ' + currentTimeScale , sliderValueWeek , 0, weekSliderMax, 1 , this._updateTimeSliderWeeks.bind(this), false, 250 ) ;

      } else if ( currentTimeScale === 'Years' ) {
        let sliderValue = this.state.gridData.allYearsStringArray[sliderValueYear];
        activeSlider = createSlider('Slide to adjust ' + currentTimeScale, sliderValue , 0, this.state.gridData.allYearsStringArray.length -1, 1 , this._updateTimeSliderPeriods.bind(this), false, 250) ;

      } else if ( currentTimeScale === 'Months' ) {
        let sliderValue = this.state.gridData.allMonthsStringArray[sliderValueMonth];
        activeSlider = createSlider('Slide to adjust ' + currentTimeScale, sliderValue , 0, this.state.gridData.allMonthsStringArray.length -1, 1 , this._updateTimeSliderPeriods.bind(this), false, 250) ;

      } else if ( currentTimeScale === 'WeekNo' ) {
        let sliderValue = this.state.gridData.allWeekNosStringArray[sliderValueWeekNo];
        activeSlider = createSlider('Slide to adjust ' + currentTimeScale, sliderValue , 0, this.state.gridData.allWeekNosStringArray.length -1, 1 , this._updateTimeSliderPeriods.bind(this), false, 250) ;

      }

      timeSlider = 
          <div className={ styles.timeSlide } style={{ }} onClick={ this._updateCurrentTimeScale.bind(this) }>
            { activeSlider }
          </div>;
    } 
    
    let legendSample = { width: '15px', height: '15px', marginTop: '5px' };
    let spacerLegend = { width: '15px', height: '15px', marginTop: '5px', border: '1px dotted gray' };

    let legend = <div className={styles.legend} >
      <div><div>{ this.state.gridData.dataLevelLabels[0] }</div><div style={ {...dataLevel0Style,...legendSample} } ></div></div>
      <div><div>{ this.state.gridData.dataLevelLabels[1] }</div><div style={ {...dataLevel1Style,...legendSample} } ></div></div>
      <div><div>{ this.state.gridData.dataLevelLabels[2] }</div><div style={ {...dataLevel2Style,...legendSample} } ></div></div>
      <div><div>{ this.state.gridData.dataLevelLabels[3] }</div><div style={ {...dataLevel3Style,...legendSample} } ></div></div>
      <div><div>Space</div><div style={ {...dataLevelMinus1Style,...spacerLegend} } ></div></div>
    </div>;

    const months : any[] = this.state.monthLables;
    const days : any[] = weekday3['en-us'];

    //let fillerDays = this.props.monthGap === "2" ? 14 : this.props.monthGap === "1" ? 7 : 0 ;
    let monthGap = parseInt( this.props.gridStyles.monthGap ) * 2;
    const gridTemplateColumns : string = this.state.monthScales.map( v => 20* (v + monthGap ) *.9 + 'px').join( ' ');

    /**
     * months were:   monthLabels
     *         <ul className={ styles.months } style={{ listStyleType: 'none', gridTemplateColumns: gridTemplateColumns, transform: sliderTransform, transition: 'transform .3s cubic-bezier(0, .52, 0, 1)' }}>
     * 
     *                { months.map( m=> { return <li> { m } </li> ; }) }  
     */

    console.log( 'yearLabels: ', yearLabels );
    let theGraph = <div className={styles.graph} style={{ width: '900px' }}>
        <ul className={ styles.years } style={{ listStyleType: 'none', }}>
          { yearLabels.map( m=> { return <li> { m } </li> ; }) }
        </ul>
        <ul className={ styles.months } style={{ listStyleType: 'none', }}>
          { monthLabels.map( m=> { return <li> { m } </li> ; }) }
        </ul>
        <ul className={styles.days} style={{ listStyleType: 'none' }}>
            { days.map( d=> { return <li> { d } </li> ; }) }
        </ul>
        { gridElement }
        <div className={ styles.graphFooter }>
          { metrics } { timeSlider } { legend }
        </div>
      </div>;

    if ( this.state.errMessage !== '' && this.state.errMessage != null ) {
      let errMessageString : any = this.state.errMessage;
      let extraMessage1 = errMessageString.indexOf('Error making HttpClient request in queryable [404]') > -1 ? 'Verify Web URL is correct': null ;
      let extraMessage2 = errMessageString.indexOf('Error making HttpClient request in queryable [404]') > -1 ? this.props.parentListWeb.replace( this.props.tenant, '' ) : null ;

      theGraph = <div style={{ textAlign: 'center', margin: '50px', height: '100px', width: '80%%'}}>
                    <span style={{ fontSize: 'larger', fontWeight: 600, paddingTop: '40px'}}>
                      <mark>{ this.state.errMessage }</mark>
                    </span><p style={{ fontSize: 'larger', fontWeight: 600 }}> { extraMessage1 } : { extraMessage2 } </p></div> ;
    } else if ( this.state.allLoaded === true && this.state.searchedItems && this.state.searchedItems.length === 0 ) {
          theGraph = <div style={{ textAlign: 'center', margin: '50px', height: '100px', width: '80%'}}>
                    <span style={{ fontSize: 'larger', fontWeight: 600, paddingTop: '40px'}}>
                      Sorry but there were no items found meeting your search criteria!
                    </span></div> ;
    }



    let panel = null;
    
    if ( this.state.showItems > -1 ) {
      let panelItems: any[] = this.state.gridData.allDataPoints[ this.state.showItems ].items;
      panel =     
      <div><Panel
        isOpen={ this.state.showItems === -1 ? false : true }
        // this prop makes the panel non-modal
        isBlocking={true}
        onDismiss={ this._onCloseItems.bind(this) }
        closeButtonAriaLabel="Close"
        type = { PanelType.large }
        isLightDismiss = { true }
        >
          <EsItems 
            pickedWeb  = { this.props.pickedWeb }
            pickedList = { this.props.pickedList }
            theSite = {null }
  
            items = { panelItems }
            itemsAreDups = { false }
            itemsAreFolders = { false }

            duplicateInfo = { null }
            heading = { `${ this.props.esItemsHeading }` }
            // batches = { batches }
            icons = { [ ]}
            emptyItemsElements = { ['Nothing was found'] }
                          
            dataOptions = { this.props.dataOptions }
            uiOptions = { this.props.uiOptions }

            sharedItems = { [] }
            
            itemType = { 'Items' }
            
          ></EsItems>
      </Panel></div>;
  
    }

    return (
      <div className={ styles.gridcharts }>
        <div className={ styles.container }>
          { searchStack }
          { theGraph }
          { panel }
        </div>
      </div>
    );
  }

  private _showTheseItems( label: string, index: number, ev: any ) {

    this.setState({
      showItems: index,
    });
  }


  private _onCloseItems( event ){
    this.setState({
      showItems: -1,
    });
  }

  private _toggleInfoPages() {
    this.setState({
      showTips: !this.state.showTips,
    });

  }


/***
 *         db    db d8888b.      .d8888. db      d888888b d8888b. d88888b d8888b. 
 *         88    88 88  `8D      88'  YP 88        `88'   88  `8D 88'     88  `8D 
 *         88    88 88oodD'      `8bo.   88         88    88   88 88ooooo 88oobY' 
 *         88    88 88~~~          `Y8b. 88         88    88   88 88~~~~~ 88`8b   
 *         88b  d88 88           db   8D 88booo.   .88.   88  .8D 88.     88 `88. 
 *         ~Y8888P' 88           `8888Y' Y88888P Y888888P Y8888D' Y88888P 88   YD 
 *                                                                                
 *                                                                                
 */
  
private _updateTimeSliderWeeks(newValue: number){

  let now = new Date();
  let then = new Date();
  then.setMinutes(then.getMinutes() + newValue);

  if ( this.props.scaleMethod === 'slider' || this.props.scaleMethod === 'blink' ) {
    //Just update slider, render method does transition with css
    this.setState({
      sliderValueWeek: newValue,
    });
  } else if ( this.props.scaleMethod === 'TBD' ) { //Update grid selected elements and date range

  }

}

private _updateCurrentTimeScale( e: any ) {
  let currentTimeScale : ITimeScale = this.state.currentTimeScale;

  if ( e.ctrlKey === true ) {
    console.log('_updateCurrentTimeScale CTRL clicked');
    if ( currentTimeScale === 'Weeks' ) { currentTimeScale = 'Years' ; }
    else if ( currentTimeScale === 'Years' ) { currentTimeScale = 'Months' ; }
    else if ( currentTimeScale === 'Months' ) { currentTimeScale = 'WeekNo' ; }
    else if ( currentTimeScale === 'WeekNo' ) { currentTimeScale = 'Weeks' ; }
  
    this.setState({
      currentTimeScale: currentTimeScale,
    });

  }


}

private _updateTimeSliderPeriods(newValue: number){
  let currentTimeScale : ITimeScale = this.state.currentTimeScale;
  let now = new Date();
  let then = new Date();
  then.setMinutes(then.getMinutes() + newValue);

  if ( this.props.scaleMethod === 'slider' || this.props.scaleMethod === 'blink' ) {
    //Just update slider, render method does transition with css

    if ( currentTimeScale === 'Weeks' ) { this.setState({ sliderValueWeek: newValue, }) ; }
    else if ( currentTimeScale === 'Years' ) { this.setState({ sliderValueYear: newValue, }) ; }
    else if ( currentTimeScale === 'Months' ) { this.setState({ sliderValueMonth: newValue, }) ; }
    else if ( currentTimeScale === 'WeekNo' ) { this.setState({ sliderValueWeekNo: newValue, }) ; }

  } else if ( this.props.scaleMethod === 'TBD' ) { //Update grid selected elements and date range

  }


}



private _updateChoiceSlider(newValue: number){

  let choiceSliderDropdown = this.state.choiceSliderDropdown;


  let theChoice = newValue > -1 ? `${ this.state.dropDownItems[choiceSliderDropdown][newValue].text }` : '' ;
  console.log('_updateChoiceSlider: choiceSliderDropdown, newValue, theChoice', choiceSliderDropdown, newValue, theChoice );

  this.setState({
    choiceSliderValue: newValue,
  });

  this.fullSearch( theChoice, choiceSliderDropdown, null, this.state.currentTimeScale );

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

 /**
  * Based on PivotTiles.tsx
  * @param item
  */
  private textSearch = ( searchText: string ): void => {

    this.fullSearch( null, null, searchText, this.state.currentTimeScale );

  }


  private changeOperation ( operation: any , ev: any ) {

    let gridData : IGridchartsData = this.buildGridData (this.state.allItems, operation, this.state.gridData.sizeOrCount );

    gridData= this.buildVisibleItems ( gridData );

    const s1 = gridData.startDate.getMonth();
    const s2 = s1 + 12;

    const monthLables = monthStr3["en-us"].concat( ... monthStr3["en-us"] ).slice(s1,s2) ;
    const monthScales = [ 4,4,4,5,4,4,5,4,4,5,4,5   ,   4,4,4,5,4,4,5,4,4,5,4,5 ].slice(s1,s2) ;

    this.setState({
      valueOperator: operation,
      gridData: gridData,

      choiceSliderDropdown: null,
      showChoiceSlider: false,
      dropdownColumnIndex: null,
      choiceSliderValue: null,

      searchedItems: this.state.allItems, //newFilteredItems,  //Replaced with theseItems to update when props change.
      searchCount: this.state.allItems.length,
      searchText: '',
      searchMeta: [],

      monthLables: monthLables,
      monthScales: monthScales,

    });

  }

  public searchForItems = (item, choiceSliderDropdown: number, DDLabel: string, ev: any): void => {

    let choiceSliderValue = null;  //choiceSliderValue

    let showChoiceSlider = this.state.showChoiceSlider;
    if ( ev.ctrlKey === true ) { 
      showChoiceSlider = true;
    } else if ( ev.altKey === true ) { 
      showChoiceSlider = false;
    }

    this.state.dropDownItems[choiceSliderDropdown].map( ( dd, ddIndex ) => {
      if ( dd.text === item ) { choiceSliderValue = ddIndex ; }
    });

    this.setState({
      choiceSliderDropdown: choiceSliderDropdown, //Number of Dropdown ( ie 1 2 or 3 )
      choiceSliderValue: choiceSliderValue, // Selected Choice of Dropdown
      showChoiceSlider: showChoiceSlider,
    });

    item = this.state.fetchList.dropDownColumns[ choiceSliderDropdown ] + '|>|' + item;
    console.log('searchForItems: ',item, choiceSliderDropdown, choiceSliderValue, ev ) ;
    this.fullSearch( item, choiceSliderDropdown, null, this.state.currentTimeScale );

  }

  public fullSearch = (item: any, choiceSliderDropdown: number, searchText: string , currentTimeScale: ITimeScale, ): void => {

    //This sends back the correct pivot category which matches the category on the tile.
    let e: any = event;

    /*
    console.log('searchForItems: e',e);
    console.log('searchForItems: item', item);
    console.log('searchForItems: this', this);


   
   if ( currentTimeScale === 'Weeks' ) { this.setState({ sliderValueWeek: newValue, }) ; }
   else if ( currentTimeScale === 'Years' ) { this.setState({ sliderValueYear: newValue, }) ; }
   else if ( currentTimeScale === 'Months' ) { this.setState({ sliderValueMonth: newValue, }) ; }
   else if ( currentTimeScale === 'WeekNo' ) { this.setState({ sliderValueWeekNo: newValue, }) ; }
    */

    let searchItems : IGridItemInfo[] = [];
    let newFilteredItems: IGridItemInfo[]  = [];

    searchItems =this.state.allItems;

    let searchCount = searchItems.length;

    let selectedDropdowns = this.state.selectedDropdowns;
    let dropDownItems = this.state.dropDownItems;

    if ( choiceSliderDropdown !== null && choiceSliderDropdown !== undefined && choiceSliderDropdown > -1 ) { //Then this is a choice dropdown filter
      //This map updates the dropdowns to find the selected one.
      let dropDownLabel = item.indexOf('|>|') > -1 ? item.split('|>|')[1] : item;
      dropDownItems[choiceSliderDropdown].map( thisChoice => {
        if ( choiceSliderDropdown === null && thisChoice.text === dropDownLabel ) { thisChoice.isSelected = true ; }  else { thisChoice.isSelected = false;} 
      });

      //This map updates the array of selected dropdowns
      selectedDropdowns.map( (dd, index ) => {
        selectedDropdowns[index] = choiceSliderDropdown === index ? dropDownLabel : ''; 
      });

      if ( dropDownLabel === '' ) {
        newFilteredItems = searchItems;
      } else {
        //This loop actually finds the newFilteredItems
        for (let thisItem of searchItems) {
          let searchChoices = thisItem.meta ;
          if(searchChoices.indexOf( item ) > -1) {
            //console.log('fileName', fileName);
            newFilteredItems.push(thisItem);
          }
        }
      }
    } else { //This is a text box filter

      //Clears the selectedDropdowns array
      selectedDropdowns.map( (dd, index ) => {
          selectedDropdowns[index] = ''; 
      });

      //Sets isSelected on all dropdown options to false
      dropDownItems.map ( ( thisDropDown ) => {
        thisDropDown.map( thisChoice => {
         thisChoice.isSelected = false;
        });
      });

      if ( searchText == null || searchText === '' ) {
        newFilteredItems = searchItems;
      } else {
        let searchTextLC = searchText.toLowerCase();
        for (let thisItem of searchItems) {
          if(thisItem.searchString.indexOf( searchTextLC ) > -1) {
            newFilteredItems.push(thisItem);
          }
        }
      }
    }

    searchCount = newFilteredItems.length;

    let gridData : IGridchartsData = this.buildGridData ( newFilteredItems, this.state.valueOperator, this.state.gridData.sizeOrCount );
    
    const s1 = gridData.startDate.getMonth();
    const s2 = s1 + 12;

    const monthLables = monthStr3["en-us"].concat( ... monthStr3["en-us"] ).slice(s1,s2) ;
    const monthScales = [ 4,4,4,5,4,4,5,4,4,5,4,5   ,   4,4,4,5,4,4,5,4,4,5,4,5 ].slice(s1,s2) ;

    this.setState({
      /*          */
        searchedItems: newFilteredItems, //newFilteredItems,  //Replaced with theseItems to update when props change.
        searchCount: newFilteredItems.length,
        searchText: searchText,
        searchMeta: [],
        dropDownItems: dropDownItems,
        selectedDropdowns: selectedDropdowns,
        dropdownColumnIndex: choiceSliderDropdown,
        gridData: gridData,
        allLoaded: true,
        monthLables: monthLables,
        monthScales: monthScales,
        lastStateChange: 'searchForItems: ' + item,
        sliderValueWeek: 0,

    });

    return ;
    
  }

  /***
 *     .d8b.  d8888b. d8888b.      d888888b d888888b d88888b .88b  d88. .d8888.      d888888b  .d88b.       .d8888. d888888b  .d8b.  d888888b d88888b 
 *    d8' `8b 88  `8D 88  `8D        `88'   `~~88~~' 88'     88'YbdP`88 88'  YP      `~~88~~' .8P  Y8.      88'  YP `~~88~~' d8' `8b `~~88~~' 88'     
 *    88ooo88 88   88 88   88         88       88    88ooooo 88  88  88 `8bo.           88    88    88      `8bo.      88    88ooo88    88    88ooooo 
 *    88~~~88 88   88 88   88         88       88    88~~~~~ 88  88  88   `Y8b.         88    88    88        `Y8b.    88    88~~~88    88    88~~~~~ 
 *    88   88 88  .8D 88  .8D        .88.      88    88.     88  88  88 db   8D         88    `8b  d8'      db   8D    88    88   88    88    88.     
 *    YP   YP Y8888D' Y8888D'      Y888888P    YP    Y88888P YP  YP  YP `8888Y'         YP     `Y88P'       `8888Y'    YP    YP   YP    YP    Y88888P 
 *                                                                                                                                                    
 *                                                                                                                                                    
 */


  private addTheseItemsToState( fetchList: IGridList, theseItems , errMessage : string, allNewData : boolean = true ) {

      if ( theseItems.length < 300 ) {
          console.log('addTheseItemsToState theseItems: ', theseItems);
      } {
          console.log('addTheseItemsToState theseItems: QTY: ', theseItems.length );
      }

      let allItems = allNewData === false ? this.state.allItems : theseItems;

      let gridData : IGridchartsData = this.buildGridData (theseItems, this.state.valueOperator, this.state.gridData.sizeOrCount );

      gridData= this.buildVisibleItems ( gridData );

      let dropDownItems : IDropdownOption[][] = allNewData === true ? this.buildDataDropdownItems( allItems ) : this.state.dropDownItems ;
      
      const s1 = gridData.startDate.getMonth();
      const s2 = s1 + 12;

      const monthLables = monthStr3["en-us"].concat( ... monthStr3["en-us"] ).slice(s1,s2) ;
      const monthScales = [ 4,4,4,5,4,4,5,4,4,5,4,5   ,   4,4,4,5,4,4,5,4,4,5,4,5 ].slice(s1,s2) ;

      this.setState({
        /*          */
          allItems: allItems,
          searchedItems: theseItems, //newFilteredItems,  //Replaced with theseItems to update when props change.
          searchCount: theseItems.length,
          dropDownItems: dropDownItems,
          errMessage: errMessage,
          searchText: '',
          searchMeta: [],
          fetchList: fetchList,
          gridData: gridData,
          allLoaded: true,
          monthLables: monthLables,
          monthScales: monthScales,

      });

      console.log('loadedState:', this.state );
      //This is required so that the old list items are removed and it's re-rendered.
      //If you do not re-run it, the old list items will remain and new results get added to the list.
      //However the list will show correctly if you click on a pivot.
      //this.searchForItems( '', this.state.searchMeta, 0, 'meta' );
      return true;
  }

  private buildVisibleItems( gridData : IGridchartsData ) {

    return gridData;
  }


  private buildDataDropdownItems( allItems : IGridItemInfo[] ) {

    let dropDownItems : IDropdownOption[][] = [];

    this.props.columns.dropDownColumns.map( ( col, colIndex ) => {

      let actualColName = col.replace('>', '' ).replace('+', '' ).replace('-', '' );
      let parentColName = colIndex > 0 && col.indexOf('>') > -1 ? this.props.columns.dropDownColumns[colIndex - 1] : null;
      parentColName = parentColName !== null ? parentColName.replace('>', '' ).replace('+', '' ).replace('-', '' ) : null;

      let thisColumnChoices : IDropdownOption[] = [];
      let foundChoices : string[] = [];
      allItems.map( item => {
        let thisItemsChoices = item[ actualColName ];
        if ( actualColName.indexOf( '/') > -1 ) {
          let parts = actualColName.split('/');
          thisItemsChoices = item[ parts[0] ] ? item[ parts[0] ] [parts[1]] :  `. missing ${ parts[0] }`;
        }
        if ( parentColName !== null ) { thisItemsChoices = item[ parentColName ] + ' > ' + item[ actualColName ] ; }
        if ( thisItemsChoices && thisItemsChoices.length > 0 ) {
          if ( foundChoices.indexOf( thisItemsChoices ) < 0 ) {
            if ( thisColumnChoices.length === 0 ) { thisColumnChoices.push( { key: '', text: '- all -' } ) ; }
            thisColumnChoices.push( { key: thisItemsChoices, text: thisItemsChoices } ) ;
            foundChoices.push( thisItemsChoices ) ;
          }
        }
      });

      dropDownItems.push( thisColumnChoices ) ;

    });

    return dropDownItems;

  }



/***
 *    d8888b. db    db d888888b db      d8888b.       d888b  d8888b. d888888b d8888b.      d8888b.  .d8b.  d888888b  .d8b.  
 *    88  `8D 88    88   `88'   88      88  `8D      88' Y8b 88  `8D   `88'   88  `8D      88  `8D d8' `8b `~~88~~' d8' `8b 
 *    88oooY' 88    88    88    88      88   88      88      88oobY'    88    88   88      88   88 88ooo88    88    88ooo88 
 *    88~~~b. 88    88    88    88      88   88      88  ooo 88`8b      88    88   88      88   88 88~~~88    88    88~~~88 
 *    88   8D 88b  d88   .88.   88booo. 88  .8D      88. ~8~ 88 `88.   .88.   88  .8D      88  .8D 88   88    88    88   88 
 *    Y8888P' ~Y8888P' Y888888P Y88888P Y8888D'       Y888P  88   YD Y888888P Y8888D'      Y8888D' YP   YP    YP    YP   YP 
 *                                                                                                                          
 *                                                                                                                          
 */

  private buildGridData ( allItems : IGridItemInfo[], valueOperator: IValueOperator, sizeOrCount: ISizeOrCount ) {

    let count = allItems.length;

    let allDateArray : any[] = [];
    let allDateStringArray : string[] = [];

    let allYearsStringArray: string[] = [];
    let allMonthsStringArray: string[] = [];
    let allWeekNosStringArray: string[] = [];

    let allDataPoints : IGridchartsDataPoint[] = [];

    /**
     * Get entire date range
     * miliseconds for "2021-01-31" is 1612127321000
     * 
     * 1012127321000; 
     * 1612127321000
     */

    let firstTime = 2512127321000; 
    let lastTime = 1012127321000;
    let firstDate = "";
    let lastDate = "";

    allItems.map( item => {
      let theStartTimeMS = item['time' + this.props.columns.dateColumn ].milliseconds;
      let theStartTimeStr = item['time' + this.props.columns.dateColumn ].theTime;

      if ( theStartTimeMS > lastTime ) { 
        lastTime = theStartTimeMS ; 
        lastDate = theStartTimeStr ; }

      if ( theStartTimeMS < firstTime ) { 
        firstTime = theStartTimeMS ; 
        firstDate = theStartTimeStr ; }

    });

    let startDate = new Date( firstDate );
    // let gridStart = getOffSetDayOfWeek( firstDate, 7, 'prior' ); //This gets prior sunday
    let gridStart  = new Date( startDate.getFullYear(), startDate.getMonth() , 1 ); //First day of this month

    let priorSundayStart = getOffSetDayOfWeek( gridStart.toDateString(), 7, 'prior' ); //This gets prior sunday
    
    let leadingBlanks = getTimeDelta( priorSundayStart, gridStart, 'days' ) + 1; //Days gives full days but not difference between dates so I'm taking away 1 day.

    gridStart.setHours(0,0,0,0);
    let endDate = getOffSetDayOfWeek( lastDate, 7, 'next' );
    endDate.setHours(0,0,0,0);

    // Last day of current month: https://stackoverflow.com/a/222439
    let gridEnd  = new Date( endDate.getFullYear(), endDate.getMonth() + 1, 0 );
    //let gridEnd = new Date( tempEnd.toLocaleString() );
    allDateArray = this.getDates( gridStart, gridEnd);
    allDateArray.map ( d => { 
      allDateStringArray.push( d.toLocaleDateString() ) ;

      let thisYear = d.getFullYear();
      let yearMonth : any = getYearMonthLabel(d);
      let yearWeek : any = getYearWeekLabel(d);

      if (  allYearsStringArray.indexOf( thisYear.toString() ) < 0 ) {  allYearsStringArray.push( thisYear.toString() ) ; }
      if (  allMonthsStringArray.indexOf( yearMonth ) < 0 ) {  allMonthsStringArray.push( yearMonth ) ; }
      if (  allWeekNosStringArray.indexOf( yearWeek ) < 0 ) {  allWeekNosStringArray.push( yearWeek ) ; }

    });

    /**
     * Build the IGridchartsDataPoint[] array
     */

    allDateArray.map( theDate => {
      allDataPoints.push( {
        date: theDate,

        dateNo: theDate.getDate(),
        dayNo: theDate.getDay(),
        week: null,
        month: theDate.getMonth(),
        year: theDate.getFullYear(),
        yearMonth: getYearMonthLabel( theDate ),
        yearWeek: getYearWeekLabel( theDate ),

        yearIndex: null,
        yearMonthIndex: null,
        yearWeekIndex: null,

        dateString: theDate.toLocaleDateString(),

        label: '',
        dataLevel: null,
        value: null,
        Count: 0,
        Sum: null,
        Avg: null,
        Min: null,
        Max: null,
        values: [],
        valuesString: [],
        items: [],
      });
    });

    /**
     * Go through items and add to allDataPoints
     */

    let minValue = 951212732100099;
    let maxValue = -951212732100099;
    let gridDataTotal = 0;
    
    allItems.map( item => {
      let itemDateProp = item['time' + this.props.columns.dateColumn ];
      let itemDateDate = new Date( itemDateProp.theTime );
      let itemDate = itemDateDate.toLocaleDateString();
      let dateIndex = allDateStringArray.indexOf( itemDate ) ;
      item.dateIndex = dateIndex;

      item.dateNo = itemDateProp.date;
      item.dayNo = itemDateProp.day;
      item.week = itemDateProp.week;
      item.month = itemDateProp.month;
      item.year = itemDateProp.year;
      
      let yearMonth : any =getYearMonthLabel( itemDateDate ) ;
      let yearWeek : any = getYearWeekLabel( itemDateDate ) ;

      item.yearMonth = yearMonth;
      item.yearWeek = yearWeek;

      item.yearIndex = allYearsStringArray.indexOf( item.year.toString() ) ;
      item.yearMonthIndex = allMonthsStringArray.indexOf( yearMonth ) ;
      item.yearWeekIndex = allWeekNosStringArray.indexOf( yearWeek ) ;

      item.meta.push( item.yearMonth ) ;
      item.meta.push( item.yearWeek ) ;
      item.meta.push( item.year.toString() ) ;

      item.searchString += 'yearMonth=' + item.yearMonth + '|||' + 'yearWeek=' + item.yearWeek + '|||' + 'year=' + item.year + '|||' + 'week=' + item.week + '|||';

      //Copied section from GridCharts VVVV
      let valueColumn = item[ this.props.columns.valueColumn ];
      let valueType = typeof valueColumn;

      if ( valueType === 'string' ) { valueColumn = parseFloat( valueColumn ) ; }
      else if ( valueType === 'number' ) { valueColumn = parseFloat( valueColumn ) ; }
      else if ( valueType === 'boolean' ) { valueColumn = valueColumn === true ? 1 : 0 ; }
      else if ( valueType === 'object' ) { valueColumn = 0 ; }
      else if ( valueType === 'undefined' ) { valueColumn = 0 ; }
      else if ( valueType === 'function' ) { valueColumn = 0 ; }
      //Copied section from GridCharts ^^^^
      
      allDataPoints[dateIndex].items.push( item );
      allDataPoints[dateIndex].values.push( valueColumn );
      allDataPoints[dateIndex].valuesString.push( valueColumn.toFixed(2) );

      allDataPoints[dateIndex].Count ++;
      allDataPoints[dateIndex].Sum += valueColumn;      
      if ( allDataPoints[dateIndex].Min === null || allDataPoints[dateIndex].Min > valueColumn ) { allDataPoints[dateIndex].Min = valueColumn; }  
      if ( allDataPoints[dateIndex].Max === null || allDataPoints[dateIndex].Max < valueColumn ) { allDataPoints[dateIndex].Max = valueColumn; }  

      if ( allDataPoints[dateIndex].yearIndex === null ) { allDataPoints[dateIndex].yearIndex = item.yearIndex; }  
      if ( allDataPoints[dateIndex].yearMonthIndex === null ) { allDataPoints[dateIndex].yearMonthIndex = item.yearMonthIndex; }  
      if ( allDataPoints[dateIndex].yearWeekIndex === null ) { allDataPoints[dateIndex].yearWeekIndex = item.yearWeekIndex; }  

      let compareValue = allDataPoints[dateIndex][ valueOperator ] ;
      if ( compareValue < minValue ) { minValue = compareValue; }
      if ( compareValue > maxValue ) { maxValue = compareValue; } 

      if ( valueOperator === 'Sum' || valueOperator === 'Avg' ) { gridDataTotal += valueColumn ; } 
      else if ( valueOperator === 'Count' ) { gridDataTotal ++ ; } 
      else if ( valueOperator === 'Max' && valueColumn > gridDataTotal ) { gridDataTotal = valueColumn ; } 
      else if ( valueOperator === 'Min' && valueColumn < gridDataTotal ) { gridDataTotal = valueColumn ; } 

    });

    if ( valueOperator === 'Avg' ) { 
      if ( count != 0 ) {
        gridDataTotal = gridDataTotal / count ;
        minValue = 951212732100099;
        maxValue = -951212732100099;
        allDataPoints.map( item => {
          if ( item.Count && item.Count !== 0 ) {
            item.Avg = item.Sum / item.Count;
            let compareValue = item.Avg;
            if ( compareValue < minValue ) { minValue = compareValue; }
            if ( compareValue > maxValue ) { maxValue = compareValue; } 
          } else {

          }
        });
      }    
    } 

    /**
     * Update datalevel based on min/max
     */
    
    let dataLevelIncriment = ( maxValue - minValue ) / 3;

    let dataLevel3: any = maxValue - 1 * dataLevelIncriment;
    let dataLevel2: any = maxValue - 2 * dataLevelIncriment;
    let dataLevel1: any = minValue;
    let dataLevelTop: any = maxValue;

    allDataPoints.map( (data, index)  => {
      data.Avg = data.Count !== null && data.Count !== undefined && data.Count !== 0 ? data.Sum / data.Count : null;
      data.value = data[ valueOperator ] ;

      if ( data.Count === 0 ) { data.dataLevel = 0 ; }
      else if ( data.value > ( maxValue - 1 * dataLevelIncriment ) ) { data.dataLevel = 3 ; }
      else if ( data.value > dataLevel2 ) { data.dataLevel = 2 ; }
      else if ( data.value >= minValue ) { data.dataLevel = 1 ; }
      else { data.dataLevel = 0 ; }
      let value = valueOperator !== 'Count' && sizeOrCount === 'size' ? getSizeLabel( data.value, 2 ) : getCountLabel( data.value );
      data.label = data.Count === 0 ? `${data.dateString} : No data available` : `${data.dateString} : ${valueOperator} = ${ value } 
        count: ${ data.items.length } idx: ${ index }`;
    });


    let totalLabel = null;
    if ( valueOperator === 'Count' ) {
      totalLabel = getCommaSepLabel( gridDataTotal );
      dataLevel3 = getCommaSepLabel( dataLevel3 ) ;
      dataLevel2 = getCommaSepLabel( dataLevel2 ) ;
      dataLevel1 = getCommaSepLabel( dataLevel1 ) ;
      dataLevelTop = getCommaSepLabel( dataLevelTop ) ;

    } else {
      totalLabel = sizeOrCount === 'size' ? getSizeLabel( gridDataTotal ) : getCountLabel( gridDataTotal );
      dataLevel3 = sizeOrCount === 'size' ? getSizeLabel( dataLevel3 ) : getCountLabel( dataLevel3 ) ;
      dataLevel2 = sizeOrCount === 'size' ? getSizeLabel( dataLevel2 ) : getCountLabel( dataLevel2 ) ;
      dataLevel1 = sizeOrCount === 'size' ? getSizeLabel( dataLevel1 ) : getCountLabel( dataLevel1 ) ;
      dataLevelTop = sizeOrCount === 'size' ? getSizeLabel( dataLevelTop ) : getCountLabel( dataLevelTop ) ;
    }

    let dataLevelLabels : string[] = [];

    dataLevelLabels.push( 'No data'); //DataLevel 0 label
    dataLevelLabels.push( '>= ' + dataLevel1 ) ; //DataLevel 1 label
    dataLevelLabels.push( '> ' + dataLevel2 ); //DataLevel 2 label
    dataLevelLabels.push( '> ' + dataLevel3 ); //DataLevel 3 label
    dataLevelLabels.push( ); //DataLevel 4 label
    
    let gridData: IGridchartsData = {
      total: gridDataTotal,
      totalLabel: totalLabel,
      count: count,
      sizeOrCount: sizeOrCount,
      leadingBlanks: leadingBlanks,
      gridStart: startDate,
      gridEnd: gridEnd,
      startDate: startDate,
      endDate: endDate,

      allWeeks: 0,
      allDateArray: allDateArray,
      allDateStringArray: allDateStringArray,
      
      allYearsStringArray: allYearsStringArray,
      allMonthsStringArray: allMonthsStringArray,
      allWeekNosStringArray: allWeekNosStringArray,

      allDataPoints: allDataPoints,
                
      visibleDataPoints: [],
      visibleDateArray: [],
      visibleDateStringArray: [],
      visibleWeeks: 0,
      maxValue: maxValue,
      minValue: minValue,

      dataLevelLabels: dataLevelLabels,

    };

    return gridData;

  }

}
