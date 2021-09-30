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

import {    IDropdownOption,  } from "office-ui-fabric-react";


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


import { IZBasicItemInfo, } from '@mikezimm/npmfunctions/dist/Lists/IListInterfaces';
import { IMyPivCat } from '@mikezimm/npmfunctions/dist/Pivots/IzPivots';

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

import { IGridList } from './GetListData';

export type ITimeScale  = 'Weeks' | 'Years' | 'Months' | "WeekNo";

export interface IGridchartsData {

    gridStart: any;
    startDate: any;
    endDate: any;
    gridEnd: any;

    allDateArray: any[];  //Used as easy date index of entire range of data... to easily find correct item in gridData
    allDateStringArray: string[];

    allYearsStringArray: string[];
    allMonthsStringArray: string[];
    allWeekNosStringArray: string[];

    allDataPoints: IGridchartsDataPoint[]; //One IGridchartsDataPoint per date between lowest and highest date range for input data
    allWeeks: number;

    visibleDataPoints: IGridchartsDataPoint[];
    visibleDateArray: any[];  //Used as easy date index of entire range of data... to easily find correct item in gridData
    visibleDateStringArray: string[];
    visibleWeeks: number;

    count: number;
    leadingBlanks: number; //Number of empty squares... basically prior month's squares that are not visible
    total: number;

    maxValue: number;
    minValue: number;

    dataLevelLabels: string[];


}

export interface IGridchartsDataPoint {
    date: any;
    dateNo: number;
    dayNo: number;
    week: number;
    month: number;

    year: number;
    yearIndex: number;

    yearMonth: string;
    yearMonthIndex: number;

    yearWeek: string;
    yearWeekIndex: number;

    dateString: string;
    label: any;
    dataLevel: number;
    value: number;
    count: number;
    sum: number;
    avg: number;
    min: number;
    max: number;
    values: number[];
    valuesString: string[];
    items: IGridItemInfo[];
}

/***
 *    d888888b      .d8888. d888888b  .d8b.  d888888b 
 *      `88'        88'  YP `~~88~~' d8' `8b `~~88~~' 
 *       88         `8bo.      88    88ooo88    88    
 *       88           `Y8b.    88    88~~~88    88    
 *      .88.        db   8D    88    88   88    88    
 *    Y888888P      `8888Y'    YP    YP   YP    YP    
 *                                                    
 *                                                    
 */

export type IStatType = 'sum' | 'max' | 'mini' | 'range' | '';

export interface IStat {
    prop: string;
    label: string;
    type: IStatType;
    val1?: any;
    val2?: any;
    result?: string;
}



export interface IGridchartsState {

    selectedYear: number; //Used to determine selected Year Pivot
    selectedUser: any; //Used to determine filter of items ( current user or everyone )

    selectedDropdowns: string[]; //array of selected choices for dropdowns
    dropDownItems: IDropdownOption[][]; //array of array of options for selected dropdown fields

    gridData: IGridchartsData; //One IGridchartsDataPoint per date between lowest and highest date range for input data

    monthLables: string[];  //Used to build the month labels on top of the gridChart
    monthScales: number[];  //Used to space the month labels on top of the gridChart

    sliderValueWeek: number;
    sliderValueYear: number;
    sliderValueMonth: number;
    sliderValueWeekNo: number;

    timeSliderScale: ITimeScale[];
    currentTimeScale: ITimeScale;

    choiceSliderValue: number;
    dropdownColumnIndex: number;
    choiceSliderDropdown: number;

    showChoiceSlider: boolean;
    
    breadCrumb: any[];

    WebpartHeight?:  number;    //Size courtesy of https://www.netwoven.com/2018/11/13/resizing-of-spfx-react-web-parts-in-different-scenarios/
    WebpartWidth?:   number;    //Size courtesy of https://www.netwoven.com/2018/11/13/resizing-of-spfx-react-web-parts-in-different-scenarios/

    webURL?: string;
    
    allowOtherSites?: boolean; //default is local only.  Set to false to allow provisioning parts on other sites.

    allLoaded: boolean;

    bannerMessage: any;

    showTips: boolean;

    searchCount: number;

    searchText: string;
    searchMeta: string[];

    searchedItems: IGridItemInfo[];
    stats: IStat[];
    first20searchedItems: IGridItemInfo[];

    allItems: IGridItemInfo[];

//    viewType?: IViewType;

    meta: string[];

    errMessage: string | JSX.Element;

    fetchList: IGridList;

    pivotCats: IMyPivCat[][];
 
    lastStateChange: string;
    stateChanges: string[]; //Log of state changes into array

  }





/***
 *    d888888b      d8888b. d8888b. d888888b db      db           d888888b d888888b d88888b .88b  d88.      d888888b d8b   db d88888b  .d88b.  
 *      `88'        88  `8D 88  `8D   `88'   88      88             `88'   `~~88~~' 88'     88'YbdP`88        `88'   888o  88 88'     .8P  Y8. 
 *       88         88   88 88oobY'    88    88      88              88       88    88ooooo 88  88  88         88    88V8o 88 88ooo   88    88 
 *       88         88   88 88`8b      88    88      88              88       88    88~~~~~ 88  88  88         88    88 V8o88 88~~~   88    88 
 *      .88.        88  .8D 88 `88.   .88.   88booo. 88booo.        .88.      88    88.     88  88  88        .88.   88  V888 88      `8b  d8' 
 *    Y888888P      Y8888D' 88   YD Y888888P Y88888P Y88888P      Y888888P    YP    Y88888P YP  YP  YP      Y888888P VP   V8P YP       `Y88P'  
 *                                                                                                                                             
 *                                                                                                                                             
 */


export interface IGridItemInfo extends IZBasicItemInfo {
    dateIndex: number;

    year: number;
    yearIndex: number;

    yearMonth: string;
    yearMonthIndex: number;

    yearWeek: string;
    yearWeekIndex: number;

    dateNo: number;
    dayNo: number;
    week: number;
    month: number;

    meta: string[];
    searchString: string;

}