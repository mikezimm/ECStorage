import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import { PageContext } from '@microsoft/sp-page-context';

import { sp, Views, IViews, ISite } from "@pnp/sp/presets/all";

import { setPageFormatting, IFPSPage } from '@mikezimm/npmfunctions/dist/Services/DOM/FPSFormatFunctions';
import { minimizeQuickLaunch } from '@mikezimm/npmfunctions/dist/Services/DOM/quickLaunch'; //For FPS Options

import { makeid, getStringArrayFromString } from '@mikezimm/npmfunctions/dist/Services/Strings/stringServices';

import { IUser } from '@mikezimm/npmfunctions/dist/Services/Users/IUserInterfaces';

import * as strings from 'ExStorageWebPartStrings';
import ExStorage from './components/ExStorage';
import { IExStorageProps, IDataOptions, IUiOptions } from './components/IExStorageProps';

import { FPSOptionsGroup, FPSBanner2Group } from '@mikezimm/npmfunctions/dist/Services/PropPane/FPSOptionsGroup';
import { WebPartInfoGroup, JSON_Edit_Link } from '@mikezimm/npmfunctions/dist/Services/PropPane/zReusablePropPane';
import { createStyleFromString, getReactCSSFromString, ICurleyBraceCheck } from '@mikezimm/npmfunctions/dist/Services/PropPane/StringToReactCSS';


import * as links from '@mikezimm/npmfunctions/dist/Links/LinksRepos';

import { IWebpartBannerProps, IWebpartBannerState } from './components/HelpPanel/banner/onNpm/bannerProps';

require('../../services/GrayPropPaneAccordions.css');

export interface IExStorageWebPartProps {
  description: string;

  // 0 - Context
  pageContext: PageContext;

  // 1 - Analytics options
  useListAnalytics: boolean;
  analyticsWeb?: string;
  analyticsList?: string;

  parentWeb: string;
  listTitle: string;
  showListDropdown: boolean;
  showSystemLists: boolean;
  excludeListTitles: string;

  //General settings for Banner Options group
  // export interface IWebpartBannerProps {
    bannerTitle: string;
    bannerStyle: string;
    showBanner: boolean;
    bannerHoverEffect: boolean;
    showTricks: boolean;

    showGoToHome: boolean;  //defaults to true
    showGoToParent: boolean;  //defaults to true
    showBannerGear: boolean;

  // }


  //General settings for FPS Options group
  searchShow: boolean;
  fpsPageStyle: string;
  fpsContainerMaxWidth: string;
  quickLaunchHide: boolean;

  uniqueId: string;

  useMediaTags: boolean;
  getSharedDetails: boolean;

  quickCloseItem: boolean;
  maxVisibleItems: number;


  /**
   * Imported for GridCharts VVVVVVVVVVVVVVVVVVV
   */
  gridColor?: 'green' | 'red' | 'blue' | 'theme';

  cellColor: string;
  yearStyles: string;
  monthStyles: string;
  dayStyles: string;
  cellStyles: string;
  cellhoverInfoColor: string;
  other: string;
  
  squareCustom: string;
  squareColor: string;
  emptyColor: string;
  backGroundColor: string;   

  monthGap: string; 
  
  otherStyles: string;
  /**
   * END Imported for GridCharts ^^^^^^^^^^^^^^^^^^^^^
   */

} 

export default class ExStorageWebPart extends BaseClientSideWebPart<IExStorageWebPartProps> {

  private currentSite: string = window.location.href;
  private minQuickLaunch: boolean = false;
  private fpsPageDone: boolean = false;
  private fpsPageArray: any[] = null;
  private currentUser: IUser = null;
  private urlVars : any;
  private allowOtherSites: boolean = false;
  private forceBanner = true ;
  private modifyBannerTitle = false ;
  private modifyBannerStyle = true ;

  public onInit():Promise<void> {
    return super.onInit().then(_ => {

      // other init code may be present

      this.urlVars = this.getUrlVars();
      console.log('urlVars:' , this.urlVars );

      let mess = 'onInit - ONINIT: ' + new Date().toLocaleTimeString();
      console.log(mess);

      //https://stackoverflow.com/questions/52010321/sharepoint-online-full-width-page
      if ( window.location.href &&  
        window.location.href.toLowerCase().indexOf("layouts/15/workbench.aspx") > 0 ) {
          
        if (document.getElementById("workbenchPageContent")) {
          document.getElementById("workbenchPageContent").style.maxWidth = "none";
        }
      } 

      /**
       * Set default page with using FPS Options for existing installed webparts
       */
      if ( this.properties.fpsPageStyle && this.properties.fpsPageStyle.length > 0 ) {} else { 
        this.properties.fpsPageStyle = "this.section.maxWidth=2200px" ;
      }
      
      if ( this.properties.uniqueId && this.properties.uniqueId.length > 0 ) {} else { 
        this.properties.uniqueId = makeid( 7 ) ;
      }
      //console.log('window.location',window.location);
      sp.setup({
        spfxContext: this.context
      });
    });

  }

  public getUrlVars(): {} {
    var vars = {};
    vars = location.search
    .slice(1)
    .split('&')
    .map(p => p.split('='))
    .reduce((obj, pair) => {
      const [key, value] = pair.map(decodeURIComponent);
      return ({ ...obj, [key]: value }) ;
    }, {});
    return vars;
  }

  public render(): void {

    //For FPS Options
    if ( this.properties.showBanner === undefined ) { this.properties.showBanner = true ; }
    this.setThisPageFormatting( this.properties.fpsPageStyle );
    this.setQuickLaunch( this.properties.quickLaunchHide );

    console.log('forceBanner, modifyBannerStyle, showBanner:' , this.forceBanner, this.modifyBannerStyle, this.properties.showBanner );

    //Be sure to always pass down an actual URL if the webpart prop is empty at this point.
    //If it's undefined, null or '', get current page context value
    // let parentWeb = this.properties.parentWeb && this.properties.parentWeb != '' ? this.properties.parentWeb : this.context.pageContext.web.absoluteUrl;
    let parentWeb = this.properties.parentWeb;
    if ( !parentWeb || parentWeb.length === 0 ) {
      // debugger;
      if ( this.currentSite.toLowerCase().indexOf('webpartdev') > -1 ) {
        this.allowOtherSites = true;
        parentWeb = 'https://autoliv.sharepoint.com/sites/MSLV5Generaltasks/';
      } else {
        parentWeb = this.context.pageContext.web.absoluteUrl;
        this.properties.parentWeb = this.context.pageContext.web.absoluteUrl;
      }
    }

    if ( this.urlVars.allowOtherSites === 'true' ) {
      this.allowOtherSites = true;
    }
    
    // let tenant = this.context.pageContext.web.absoluteUrl.replace(this.context.pageContext.web.serverRelativeUrl,"");
    let tenant = window.location.origin;

    // Always default to Documents library if nothing else is visible.
    let listTitle = this.properties.listTitle && this.properties.listTitle.length > 0 ? this.properties.listTitle : 'Documents';

    let dataOptions: IDataOptions = {
      useMediaTags: this.properties.useMediaTags,
      getSharedDetails: this.properties.getSharedDetails ? this.properties.getSharedDetails : true,
    };

    let uiOptions: IUiOptions = {
      quickCloseItem: this.properties.quickCloseItem,
      maxVisibleItems: this.properties.maxVisibleItems,
      showListDropdown: this.properties.showListDropdown,
      showSystemLists: this.properties.showSystemLists,
      excludeListTitles: this.properties.excludeListTitles,
    };

//  db   db d88888b db      d8888b.      d8888b.  .d8b.  d8b   db d8b   db d88888b d8888b. 
//  88   88 88'     88      88  `8D      88  `8D d8' `8b 888o  88 888o  88 88'     88  `8D 
//  88ooo88 88ooooo 88      88oodD'      88oooY' 88ooo88 88V8o 88 88V8o 88 88ooooo 88oobY' 
//  88~~~88 88~~~~~ 88      88~~~        88~~~b. 88~~~88 88 V8o88 88 V8o88 88~~~~~ 88`8b   
//  88   88 88.     88booo. 88           88   8D 88   88 88  V888 88  V888 88.     88 `88. 
//  YP   YP Y88888P Y88888P 88           Y8888P' YP   YP VP   V8P VP   V8P Y88888P 88   YD 
//                                                                                         
//                                                                                         

//UPDATE PER WEBPART
let gitHubRepo = links.gitRepoEasyStorageSmall;

let showTricks = false;
links.trickyEmails.map( getsTricks => {
  if ( this.context.pageContext.user.loginName && this.context.pageContext.user.loginName.toLowerCase().indexOf( getsTricks ) > -1 ) { showTricks = true ; }   } ); 
  if ( this.context.pageContext.user.loginName.indexOf( 'erri.scov') > -1 ){ showTricks = true ; }

let bannerTitle = this.modifyBannerTitle === true && this.properties.bannerTitle && this.properties.bannerTitle.length > 0 ? this.properties.bannerTitle : `Pivot Tiles`;
let bannerStyle: ICurleyBraceCheck = getReactCSSFromString( 'bannerStyle', this.properties.bannerStyle, {background: "#7777",fontWeight:600, fontSize: 'larger', height: '43px'} );
let showBannerGear = this.properties.showBannerGear === false ? false : true;

let anyContext: any = this.context;
console.log('_pageLayoutType:', anyContext._pageLayoutType );
console.log('pageLayoutType:', anyContext.pageLayoutType );

let bannerProps: IWebpartBannerProps = {

  pageContext: this.context.pageContext,
  panelTitle: 'Pivot Tiles webpart - Automated links and tiles',
  bannerWidth : this.domElement.clientWidth,
  showBanner: this.forceBanner === true || this.properties.showBanner !== false ? true : false,
  showTricks: showTricks,
  showBannerGear: showBannerGear,
  showGoToHome: this.properties.showGoToHome === false ? false : true,
  showGoToParent: this.properties.showGoToParent === false ? false : true,
  // onHomePage: anyContext._pageLayoutType === 'Home' ? true : false,
  onHomePage: this.context.pageContext.legacyPageContext.isWebWelcomePage === true ? true : false,
  hoverEffect: this.properties.bannerHoverEffect === false ? false : true,
  title: bannerStyle.errMessage !== '' ? bannerStyle.errMessage : bannerTitle ,
  bannerReactCSS: bannerStyle.errMessage === '' ? bannerStyle.parsed : { background: "yellow", color: "red", },
  gitHubRepo: gitHubRepo,
  farElements: [],
  nearElements: [],
  earyAccess: false,
  wideToggle: true,
};

//close #129:  This makes the maxWidth added in fps options apply to banner as well.
if ( this.properties.fpsContainerMaxWidth && this.properties.fpsContainerMaxWidth.length > 0 ) {
  bannerProps.bannerReactCSS.maxWidth = this.properties.fpsContainerMaxWidth;
}
             
//  d88888b d8b   db d8888b.      d8888b.  .d8b.  d8b   db d8b   db d88888b d8888b. 
//  88'     888o  88 88  `8D      88  `8D d8' `8b 888o  88 888o  88 88'     88  `8D 
//  88ooooo 88V8o 88 88   88      88oooY' 88ooo88 88V8o 88 88V8o 88 88ooooo 88oobY' 
//  88~~~~~ 88 V8o88 88   88      88~~~b. 88~~~88 88 V8o88 88 V8o88 88~~~~~ 88`8b   
//  88.     88  V888 88  .8D      88   8D 88   88 88  V888 88  V888 88.     88 `88. 
//  Y88888P VP   V8P Y8888D'      Y8888P' YP   YP VP   V8P VP   V8P Y88888P 88   YD 
//                                                                                  
//                                                                                  

    // // Object.keys( [] ).map( key => {
    // let errBannerMessage = '';
    // ['bannerStyle'].map( key => {

    //   let braced = addCurleyBraces( key, bannerProps[ key ] );
    //   if ( braced.parsed && braced.errMessage === '' ) {
    //     bannerProps[ key ] = braced.string;
    //     this.properties[ key ] = braced.string;

    //   } else { errBannerMessage = braced.errMessage; }

    // });

    // if ( errBannerMessage !== '' ) {
    //   bannerProps.title = errBannerMessage;
    //   bannerProps.style = `{"background": "yellow", "color": "red"}`;
    // }


//   .o88b. d8888b. d88888b  .d8b.  d888888b d88888b      d88888b db      d88888b .88b  d88. d88888b d8b   db d888888b 
//  d8P  Y8 88  `8D 88'     d8' `8b `~~88~~' 88'          88'     88      88'     88'YbdP`88 88'     888o  88 `~~88~~' 
//  8P      88oobY' 88ooooo 88ooo88    88    88ooooo      88ooooo 88      88ooooo 88  88  88 88ooooo 88V8o 88    88    
//  8b      88`8b   88~~~~~ 88~~~88    88    88~~~~~      88~~~~~ 88      88~~~~~ 88  88  88 88~~~~~ 88 V8o88    88    
//  Y8b  d8 88 `88. 88.     88   88    88    88.          88.     88booo. 88.     88  88  88 88.     88  V888    88    
//   `Y88P' 88   YD Y88888P YP   YP    YP    Y88888P      Y88888P Y88888P Y88888P YP  YP  YP Y88888P VP   V8P    YP    
//                                                                                                                     
//                                                                                                                     

    const element: React.ReactElement<IExStorageProps> = React.createElement(
      ExStorage,
      {
        // 0 - Context
        pageContext: this.context.pageContext,
        wpContext: this.context,
        tenant: tenant,
        urlVars: this.urlVars,
        bannerProps: bannerProps,
    
        //Size courtesy of https://www.netwoven.com/2018/11/13/resizing-of-spfx-react-web-parts-in-different-scenarios/
        WebpartElement: this.domElement,

        // 1 - Analytics options
        useListAnalytics: this.properties.useListAnalytics,
        analyticsList: strings.analyticsList,
        analyticsWeb: tenant + strings.analyticsWeb,

        //Size courtesy of https://www.netwoven.com/2018/11/13/resizing-of-spfx-react-web-parts-in-different-scenarios/
        WebpartHeight: this.domElement.getBoundingClientRect().height ,
        WebpartWidth:  this.domElement.getBoundingClientRect().width - 50 ,

        parentWeb: parentWeb,
        listTitle: listTitle,
    
        allowOtherSites: this.allowOtherSites, //default is local only.  Set to false to allow provisioning parts on other sites.
        pickedWeb : null,
        theSite: null,

        isLoaded: false,

        dataOptions: dataOptions,
        uiOptions: uiOptions,
    
        currentUser: this.currentUser,

        gridStyles: {
          cellColor: this.properties.cellColor ? this.properties.cellColor : 'green',
          yearStyles: this.properties.yearStyles ? this.properties.yearStyles : '',
          monthStyles: this.properties.monthStyles ? this.properties.monthStyles : '',
          dayStyles: this.properties.dayStyles ? this.properties.dayStyles : '',
          cellStyles: this.properties.cellStyles ? this.properties.cellStyles : '',
          cellhoverInfoColor: this.properties.cellhoverInfoColor ? this.properties.cellhoverInfoColor : '',
          other: this.properties.otherStyles ? this.properties.otherStyles : '',

          squareColor: this.properties.cellColor === 'swatch' && this.properties.squareColor ? this.properties.squareColor : '',
          squareCustom: this.properties.cellColor === 'custom' && this.properties.squareCustom && this.properties.squareCustom.length > 0 ? this.properties.squareCustom : 'transparent,#ebedf0,#c6e48b,#7bc96f,#196127',
          emptyColor: this.properties.cellColor === 'swatch' && this.properties.emptyColor ? this.properties.emptyColor : '',
          backGroundColor: this.properties.cellColor === 'swatch' && this.properties.backGroundColor ? this.properties.backGroundColor : '',

          monthGap: this.properties.monthGap === null || this.properties.monthGap === undefined || this.properties.monthGap === '' ? '1' : this.properties.monthGap ,
          
        },

      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          displayGroupsAsAccordion: true,
          groups: [
            WebPartInfoGroup( links.gitRepoEasyStorage, 'For analyzing extreme document libraries' ),

            {
              groupName: strings.BasicGroupName,
              isCollapsed: false ,
              groupFields: [
                PropertyPaneTextField('parentWeb', {
                  label: 'Site URL',
                  description: 'Will load current site by default',
                  disabled: this.allowOtherSites !== true ? true : false,
                }),

                PropertyPaneTextField('listTitle', {
                  label: 'Library Title'
                }),

                PropertyPaneToggle('showListDropdown', {
                  label: 'Show Lists Dropdown in webpart'
                }),

                PropertyPaneToggle('showSystemLists', {
                  label: 'Show System Lists',
                  disabled: this.properties.showListDropdown === true ? false : true,
                }),

                PropertyPaneTextField('excludeListTitles', {
                  label: 'Exclude these from dropdown',
                  disabled: this.properties.showListDropdown === true ? false : true,
                  description: 'Case SENSITIVE semi-colon (;) separated words'
                }),
                
                PropertyPaneToggle('useMediaTags', {
                  label: 'Include Media Tags in search',
                  disabled: false,
                }),
              ]
            },
            FPSBanner2Group( this.forceBanner , this.modifyBannerTitle, this.modifyBannerStyle, this.properties.showBanner, null, true ),
            FPSOptionsGroup( false, true, true, true ), // this group,
          ]
        }
      ]
    };
  }

  /**
   * Used with FPS Functions
   * @param quickLaunchHide 
   */
  private setQuickLaunch( quickLaunchHide: boolean ) {

    if ( quickLaunchHide === true && this.minQuickLaunch === false ) {
      minimizeQuickLaunch( document , quickLaunchHide );
      this.minQuickLaunch = true;
    }

  }

  /**
   * Used with FPS Functions
   * @param fpsPageStyle 
   */
  private setThisPageFormatting( fpsPageStyle: string ) {
    let fpsPage: IFPSPage = {
      Done: this.fpsPageDone,
      Style: fpsPageStyle,
      Array: this.fpsPageArray,
    };

    fpsPage = setPageFormatting( this.domElement, fpsPage );
    this.fpsPageArray = fpsPage.Array;
    this.fpsPageDone = fpsPage.Done;
  }

}
