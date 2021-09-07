import * as React from 'react';
import styles from './EcStorage.module.scss';
import { IEcStorageProps } from './IEcStorageProps';
import { IEcStorageState } from './IEcStorageState';
import { escape } from '@microsoft/sp-lodash-subset';


import { sp, Views, IViews, ISite } from "@pnp/sp/presets/all";

import ReactJson from "react-json-view";

import { IPickedWebBasic, IPickedList, }  from '@mikezimm/npmfunctions/dist/Lists/IListInterfaces';
import { IUser } from '@mikezimm/npmfunctions/dist/Services/Users/IUserInterfaces';

import { getSiteInfo, getWebInfoIncludingUnique } from '@mikezimm/npmfunctions/dist/Services/Sites/getSiteInfo';
import { cleanURL } from '@mikezimm/npmfunctions/dist/Services/Strings/urlServices';
import { getHelpfullErrorV2 } from '@mikezimm/npmfunctions/dist/Services/Logging/ErrorHandler';

export default class EcStorage extends React.Component<IEcStorageProps, IEcStorageState> {

  
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



public constructor(props:IEcStorageProps){
  super(props);

  let parentWeb = cleanURL(this.props.parentWeb);

  this.state = {

        pickedList : null,
        pickedWeb : this.props.pickedWeb,
        allLoaded: false,
      
        showPane: false,
        currentUser: null,
        isCurrentWeb: null,

        parentWeb: parentWeb,
        listTitle: this.props.listTitle,

        allowRailsOff: false,

        theSite: null,
  
  };
}


public componentDidMount() {
  if ( this.props.currentUser === null ) {
    this.getCurrentUser();
  }
  this.updateWebInfo( this.props.parentWeb );
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

  let isCurrentWeb: boolean = false;
  if ( webUrl.toLowerCase().indexOf( this.props.pageContext.web.serverRelativeUrl.toLowerCase() ) > -1 ) { isCurrentWeb = true ; }
  this.setState({ parentWeb: webUrl, stateError: stateError, pickedWeb: pickedWeb, isCurrentWeb: isCurrentWeb, theSite: theSite });

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

  public render(): React.ReactElement<IEcStorageProps> {
    return (
      <div className={ styles.ecStorage }>
        <div className={ styles.container }>

          <span className={ styles.title }>Welcome to SharePoint!</span>
          <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
          <p className={ styles.description }>{escape(this.props.parentWeb)}</p>
          <div>{ this.state.currentUser ? this.state.currentUser.Title : null }</div>

          <div style={{ overflowY: 'auto' }}>
              <ReactJson src={ this.state.currentUser } collapsed={ true } displayDataTypes={ true } displayObjectSize={ true } enableClipboard={ true } />
          </div>

        </div>
      </div>
    );
  }

  public async getCurrentUser(): Promise<void> {
    let currentUser : IUser =  null;
    await sp.web.currentUser.get().then((r) => {
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
      this.setState({ currentUser: currentUser });
    });
  
  }

}
