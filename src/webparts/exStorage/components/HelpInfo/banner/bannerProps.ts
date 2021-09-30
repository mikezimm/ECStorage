
export interface IWebpartBannerProps {
	title: string;
	style: string;
	showBanner: boolean;
	showTricks: boolean;
	
	gitHubRepo: any; // replace with IRepoLinks from npmFunctions v0.1.0.3
	
}

export interface IWebpartBannerState {
	showPanel: boolean;
	selectedKey: string;
}
