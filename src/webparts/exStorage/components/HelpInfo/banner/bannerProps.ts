import { IRepoLinks } from '@mikezimm/npmfunctions/dist/HelpInfo/Links/CreateLinks';

export interface IWebpartBannerProps {
	title: string;
	style: string;
	showBanner: boolean;
	showTricks: boolean;
	gitHubRepo: IRepoLinks; // replace with IRepoLinks from npmFunctions v0.1.0.3
	
}

export interface IWebpartBannerState {
	showPanel: boolean;
	selectedKey: string;
}
