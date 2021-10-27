import { IRepoLinks } from '@mikezimm/npmfunctions/dist/HelpInfo/Links/CreateLinks';

import { Panel, IPanelProps, PanelType } from 'office-ui-fabric-react/lib/Panel';

export interface IWebpartBannerProps {
	title: string;
	panelTitle: string;
	styleString?: string;
	bannerReactCSS?: React.CSSProperties;
	earyAccess?: boolean; //Auto add early access warning in panel
	showBanner: boolean;
	showTricks: boolean;
	gitHubRepo: IRepoLinks; // replace with IRepoLinks from npmFunctions v0.1.0.3

	toggleWide?: boolean; //enables panel width expander, true by default

	nearElements: any[];
	farElements: any[];


	
}

export interface IWebpartBannerState {
	showPanel: boolean;
	selectedKey: string;
	panelType: PanelType;
}
