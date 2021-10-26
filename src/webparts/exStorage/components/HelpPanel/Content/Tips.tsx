import * as React from 'react';

import * as fpsAppIcons from '@mikezimm/npmfunctions/dist/Icons/standardExStorage';

import { Icon  } from 'office-ui-fabric-react/lib/Icon';

const sampleIcon = <Icon iconName= { 'ExcelDocument' } style={ { fontSize: 'larger', color: 'darkgreen', padding: '0px 5px 0px 5px', } }></Icon>;

export const webParTips : any[] = [
  `CTRL-Click on some icons to auto-filter on that subject.`, 
  <span>CTRL-Click on some icons to auto-filter on that subject.</span>, 
  <span>CTRL-Click on { fpsAppIcons.CheckOutByOther } to filter on all checked out items</span>, 
  <span>CTRL-Click on { fpsAppIcons.CheckedOutByYou } to filter on all items you have checked out</span>, 
  <span>CTRL-Click on File Type Icons { sampleIcon } to filter on that file type</span>, 
  <span>In Sharing, click on <span style={{fontWeight: 600 }}>@user.name, Date, FileTypeIcon, FileName</span> to filter on that info</span>, 
];

export function getRandomTip() {

  return webParTips[Math.floor(Math.random() * webParTips.length)];

}