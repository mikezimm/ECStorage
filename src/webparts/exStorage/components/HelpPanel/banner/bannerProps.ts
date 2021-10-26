import { IRepoLinks } from '@mikezimm/npmfunctions/dist/HelpInfo/Links/CreateLinks';

export interface IWebpartBannerProps {
	title: string;
	panelTitle: string;
	styleString?: string;
	bannerReactCSS?: React.CSSProperties;
	earyAccess?: boolean; //Auto add early access warning in panel
	showBanner: boolean;
	showTricks: boolean;
	gitHubRepo: IRepoLinks; // replace with IRepoLinks from npmFunctions v0.1.0.3

	nearElements: any[];
	farElements: any[];


	
}

export interface IWebpartBannerState {
	showPanel: boolean;
	selectedKey: string;
}
