import * as React from 'react';

import * as fpsAppIcons from '@mikezimm/npmfunctions/dist/Icons/standardExStorage';

import { Icon  } from 'office-ui-fabric-react/lib/Icon';

import { buildAppWarnIcon, buildClickableIcon } from '@mikezimm/npmfunctions/dist/Icons/stdIconsBuildersV02';

import * as StdIcons from '@mikezimm/npmfunctions/dist/Icons/iconNames';

const sampleIcon = <Icon iconName= { 'ExcelDocument' } style={ { fontSize: 'larger', color: 'darkgreen', padding: '0px 5px 0px 5px', } }></Icon>;

const checkedOutToYouIcon = buildClickableIcon('eXTremeStorage', StdIcons.CheckedOutByYou , `Items you have checked out`, '#a4262c', 
  null, null, 'CheckedOutToYou' );

const checkedOutToOtherIcon = buildClickableIcon('eXTremeStorage', StdIcons.CheckedOutByOther , `Checked out by someone`, 'black', 
  null, null, 'CheckedOutByOther'); 

export const webParTips : any[] = [
  `CTRL-Click on some icons to auto-filter on that subject.`, 
  <span>CTRL-Click on some icons to auto-filter on that subject.</span>, 
  <span>CTRL-Click on <span style={{fontSize: 'larger'}}>{ checkedOutToOtherIcon }</span> to filter on all checked out items</span>, 
  <span>CTRL-Click on <span style={{fontSize: 'larger'}}>{ checkedOutToYouIcon }</span> to filter on all items you have checked out</span>, 
  <span>CTRL-Click on File Type Icons <span style={{fontSize: 'larger'}}>{ sampleIcon }</span>  to filter on that file type</span>, 
  <span>In Sharing, click on <span style={{fontWeight: 600 }}>@user.name, Date, FileTypeIcon, FileName</span> to filter on that info</span>, 
];

export function getRandomTip() {

  return webParTips[Math.floor(Math.random() * webParTips.length)];

}