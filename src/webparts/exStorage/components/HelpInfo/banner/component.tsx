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

import { IWebpartBannerProps, IWebpartBannerState } from './bannerProps';

const pivotStyles = {
	root: {
		whiteSpace: "normal",
	//   textAlign: "center"
	}};

const pivotHeading1 = 'Getting started';  //Templates
const pivotHeading2 = 'More info';  //Templates
const pivotHeading3 = 'About';  //Templates

export default class WebpartBanner extends React.Component<
	IWebpartBannerProps, IWebpartBannerState > {

    constructor(props: IWebpartBannerProps) {
			super(props);
			this.state = {
				showPanel: false,
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
	
			let GettingStartedCards: JSX.Element[] = [];
	
			GettingStartedCards.push( <AssetCard { ...assets.WelcomeToQH } /> );
			GettingStartedCards.push( <AssetCard { ...assets.NavigatQuickhelp } /> );
	
			let GettingStartedSkillpath = <AssetCard { ...assets.WelcomeToQHPlayList } />;
	
			let panelContent = <div>
				<MessageBar messageBarType={MessageBarType.warning} style={{ fontSize: 'larger' }}>
					{ `Some features highlighted in QuickHelp are not enabled at Autoliv` }
				</MessageBar>
				<h3> { `Autoliv and Quickhelp` }</h3>
				<Pivot
						// styles={ pivotStyles }
						linkFormat={PivotLinkFormat.links}
						linkSize={PivotLinkSize.normal}
						// onLinkClick={this._selectedIndex.bind(this)}
				>
						<PivotItem headerText={pivotHeading1} ariaLabel={pivotHeading1} title={pivotHeading1} itemKey={pivotHeading1} itemIcon={ 'TriangleRight12' }>
							<ol>
								<li>Click the content you want to see</li>
								<li>Enter your Autoliv email</li>
								<li>Create a profile</li>
								<li>Explore and learn new things!</li>
							</ol>
	
	
							<div className={stylesComp.bannerComponent}>
								<div className={stylesComp.flexContainer}>
									{GettingStartedCards}
								</div>
								<div className={stylesComp.flexContainer} style={{ paddingTop: '30px'}}>
									<h2>Playlist with more QuickHelp Tips</h2>
									{GettingStartedSkillpath}
								</div>
							</div>
	
						</PivotItem>
						<PivotItem headerText={pivotHeading2} ariaLabel={pivotHeading2} title={pivotHeading2} itemKey={pivotHeading2} itemIcon={ 'Info'}>
								<div style={{marginTop: '20px'}}>
									<h2>We do our best to review all content in Quickhelp</h2>
									<p>However, there will be instances where some features highlighted in quickhelp are disabled at Autoliv.</p>
									<p>If you find any of these, please help us help you by improving our content.</p>
									<p>Please submit a help ticket if you have any questions about these features.  In your incident, please include a link to the playlist or video you are referring to and the feature that you are referring to.</p>
									<p><a href="https://autolivprod.service-now.com/servicenet/" target="_blank">Submit incident in Service Now</a>
										</p>
								</div>
						</PivotItem>
						<PivotItem headerText={pivotHeading3} ariaLabel={pivotHeading3} title={pivotHeading3} itemKey={pivotHeading3} itemIcon={ null }>
								<div style={{marginTop: '20px'}}>
									<h2>Webpart verison:</h2>
									<p>{ assets.panelVersionNumber }</p>
								</div>
						</PivotItem>
				</Pivot>
			</div>;
	
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

	private _closePanel ( )  {
    this.setState({ showPanel: false,});
	}
	
	private _openPanel ( )  {
    this.setState({ showPanel: true,});
	}

}
