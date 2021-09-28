import * as React from "react";
import styles from "./banner.module.scss";
import stylesComp from "./component.module.scss";

import { escape } from "@microsoft/sp-lodash-subset";

import { Panel, IPanelProps, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { Pivot, PivotItem, IPivotItemProps, PivotLinkFormat, PivotLinkSize,} from 'office-ui-fabric-react/lib/Pivot';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';

import { Icon, IIconProps } from 'office-ui-fabric-react/lib/Icon';

import { QuichHelpVCard, AssetCard } from './AssetCard';
import * as assets from "./assets";

import WebPartLinks from './WebPartLinks';

import SinglePage from './SinglePage';
import { aboutTable } from '../Content/About';
import { devTable } from '@mikezimm/npmfunctions/dist/HelpInfo/Content/Developer';
import { gettingStartedContent } from '../Content/GettingStarted';

import { errorsContent } from '../Content/Errors';
import { advancedContent } from '../Content/Advanced';
import { futureContent } from '../Content/FuturePlans';

import { basicsContent } from '../Content/Basics';

import { tricksTable } from '../Content/Tricks';

import { IWebpartBannerProps, IWebpartBannerState } from './bannerProps';

const pivotStyles = {
	root: {
		whiteSpace: "normal",
	//   textAlign: "center"
	}};

const pivotHeading1 = 'Getting started';  //Templates
const pivotHeading2 = 'Basics';  //Templates
const pivotHeading3 = 'Advanced';  //Templates
const pivotHeading4 = 'Future';  //Templates
const pivotHeading5 = 'Dev';  //Templates
const pivotHeading6 = 'Errors';  //Templates
const pivotHeading7 = 'Tricks';  //Templates
const pivotHeading8 = 'About';  //Templates

export default class WebpartBanner extends React.Component<IWebpartBannerProps, IWebpartBannerState > {
		
    private gettingStarted= gettingStartedContent();
    private basics= basicsContent();
    private advanced= advancedContent();
    private futurePlans= futureContent();
    private dev= devTable();
		private errors= errorsContent();
		private tricks= tricksTable();
    private about= aboutTable();

    constructor(props: IWebpartBannerProps) {
			super(props);
			this.state = {
				showPanel: false,
				selectedKey: pivotHeading1,
			};
		}

		public render(): React.ReactElement<IWebpartBannerProps> {
		const { showBanner, showTricks } = this.props;
		const { showPanel } = this.state;

		if ( showBanner !== true ) {
			return (null);
		} else {


			let bannerTitle = this.props.title && this.props.title.length > 0 ? this.props.title : 'Extreme Storage Webpart';

			let classNames = [ styles.container, styles.opacity, styles.flexContainer ].join( ' ' ); //, styles.innerShadow
			let bannerContent = <div className={ classNames } style={{ height: '35px', paddingLeft: '20px', paddingRight: '20px' }}>
				<div> { bannerTitle } </div>
				<div>More information</div>
			</div>;

			let thisPage = null;

			let panelContent = null;

			if ( showPanel === true ) {
				const webPartLinks =  <WebPartLinks 
					parentListURL = { null } //Get from list item
					childListURL = { null } //Get from list item

					parentListName = { null } // Static Name of list (for URL) - used for links and determined by first returned item
					childListName = { null } // Static Name of list (for URL) - used for links and determined by first returned item

					repoObject = { this.props.gitHubRepo }
				></WebPartLinks>;

				let content = null;
				if ( this.state.selectedKey === pivotHeading1 ) {
						content = this.gettingStarted;
				} else if ( this.state.selectedKey === pivotHeading2 ) {
						content= this.basics;
				} else if ( this.state.selectedKey === pivotHeading3 ) {
						content=  this.advanced;
				} else if ( this.state.selectedKey === pivotHeading4 ) {
						content=  this.futurePlans;
				} else if ( this.state.selectedKey === pivotHeading5 ) {
						content=  this.dev;
				} else if ( this.state.selectedKey === pivotHeading6 ) {
						content=  this.errors;
				} else if ( this.state.selectedKey === pivotHeading7 ) {
						content= this.tricks;
				} else if ( this.state.selectedKey === pivotHeading8 ) {
						content= this.about;
				}

				thisPage = content === null ? null : <SinglePage 
						allLoaded={ true }
						showInfo={ true }
						content= { content }
				></SinglePage>;

				panelContent = <div>
					<MessageBar messageBarType={MessageBarType.severeWarning} style={{ fontSize: 'larger' }}>
						{ `Webpart is still under development` }
					</MessageBar>
					{ webPartLinks }
					<h3> { `Early Access webpart :)` }</h3>
					<Pivot
							// styles={ pivotStyles }
							linkFormat={PivotLinkFormat.links}
							linkSize={PivotLinkSize.normal }
							onLinkClick={this._selectedIndex.bind(this)}
					> 
						{/* { pivotItems.map( item => { return  ( item ) ; }) }
						*/}
						{ this.gettingStarted === null ? null : <PivotItem headerText={pivotHeading1} ariaLabel={pivotHeading1} title={pivotHeading1} itemKey={pivotHeading1} itemIcon={ null }/> }
						{ this.basics				 === null ? null : <PivotItem headerText={pivotHeading2} ariaLabel={pivotHeading2} title={pivotHeading2} itemKey={pivotHeading2} itemIcon={ null }/> }
						{ this.advanced			 === null ? null : <PivotItem headerText={pivotHeading3} ariaLabel={pivotHeading3} title={pivotHeading3} itemKey={pivotHeading3} itemIcon={ null }/> }
						{ this.futurePlans		 === null ? null : <PivotItem headerText={pivotHeading4} ariaLabel={pivotHeading4} title={pivotHeading4} itemKey={pivotHeading4} itemIcon={ 'RenewalFuture' }/> }
						{ this.errors 				 === null ? null : <PivotItem headerText={pivotHeading6} ariaLabel={pivotHeading6} title={pivotHeading6} itemKey={pivotHeading6} itemIcon={ 'Warning12' }/> }
						{ this.dev						 === null ? null : <PivotItem headerText={ null } ariaLabel={pivotHeading5} title={pivotHeading5} itemKey={pivotHeading5} itemIcon={ 'TestAutoSolid' }/> }
						{ this.tricks 				 === null ? null : <PivotItem headerText={ null } ariaLabel={pivotHeading7} title={pivotHeading7} itemKey={pivotHeading7} itemIcon={ 'AutoEnhanceOn' }/> }
						{ this.about 				 === null ? null : <PivotItem headerText={ null } ariaLabel={pivotHeading8} title={pivotHeading8} itemKey={pivotHeading8} itemIcon={ 'Info' }/> }
					</Pivot>
					{ thisPage }
				</div>;
			}

	
			let bannerPanel = <div><Panel
					isOpen={ showPanel }
					// this prop makes the panel non-modal
					isBlocking={true}
					onDismiss={ this._closePanel.bind(this) }
					closeButtonAriaLabel="Close"
					type = { PanelType.medium }
					isLightDismiss = { true }
				>
				{ panelContent }
			</Panel></div>;
	
			return (
				<div className={styles.bannerComponent} onClick={ this._openPanel.bind( this ) }>
					{ bannerContent }
					{ bannerPanel }
				</div>
	
			);
	
		}


	}

	public _selectedIndex = (item): void => {
    //This sends back the correct pivot category which matches the category on the tile.
    let e: any = event;

		let itemKey = item.props.itemKey;

		this.setState({ selectedKey: itemKey });
		
	}

	private _closePanel ( )  {
    this.setState({ showPanel: false,});
	}
	
	private _openPanel ( )  {
    this.setState({ showPanel: true,});
	}

}
