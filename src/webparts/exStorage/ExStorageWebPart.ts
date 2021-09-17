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
import { IExStorageProps } from './components/IExStorageProps';

import { FPSOptionsGroup } from '@mikezimm/npmfunctions/dist/Services/PropPane/FPSOptionsGroup';
import { WebPartInfoGroup, JSON_Edit_Link } from '@mikezimm/npmfunctions/dist/Services/PropPane/zReusablePropPane';
import * as links from '@mikezimm/npmfunctions/dist/HelpInfo/Links/LinksRepos';

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

  //General settings for FPS Options group
  searchShow: boolean;
  fpsPageStyle: string;
  fpsContainerMaxWidth: string;
  quickLaunchHide: boolean;

  uniqueId: string;

}

export default class ExStorageWebPart extends BaseClientSideWebPart<IExStorageWebPartProps> {

  private currentSite: string = window.location.href;
  private minQuickLaunch: boolean = false;
  private fpsPageDone: boolean = false;
  private fpsPageArray: any[] = null;
  private currentUser: IUser = null;
  private urlVars : any;
  private allowOtherSites: boolean = false;

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

    this.setThisPageFormatting( this.properties.fpsPageStyle );
    this.setQuickLaunch( this.properties.quickLaunchHide );

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


    const element: React.ReactElement<IExStorageProps> = React.createElement(
      ExStorage,
      {
        // 0 - Context
        pageContext: this.context.pageContext,
        wpContext: this.context,
        tenant: tenant,
        urlVars: this.urlVars,
    
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
        listTitle: this.properties.listTitle,
    
        allowOtherSites: this.allowOtherSites, //default is local only.  Set to false to allow provisioning parts on other sites.
        pickedWeb : null,
        theSite: null,

        isLoaded: false,
    
        currentUser: this.currentUser,

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
            FPSOptionsGroup( false, true, true, true ), // this group,
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
                  description: 'semi-colon separated words'
                }),

              ]
            }
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