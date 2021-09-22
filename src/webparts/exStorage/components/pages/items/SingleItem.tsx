import * as React from 'react';
import styles from '../../ExStorage.module.scss';
import { IEsItemsProps } from './IEsItemsProps';
import { IEsItemsState } from './IEsItemsState';
import { escape } from '@microsoft/sp-lodash-subset';


import { sp, Views, IViews, ISite } from "@pnp/sp/presets/all";
import { Web, IList, Site } from "@pnp/sp/presets/all";

import { ProgressIndicator } from 'office-ui-fabric-react/lib/ProgressIndicator';

import { Icon  } from 'office-ui-fabric-react/lib/Icon';

import { IItemDetail,  } from '../../IExStorageState';
  
const cellMaxStyle: React.CSSProperties = {
  whiteSpace: 'nowrap',
  height: '15px',
  padding: '10px 30px 0px 0px',
  fontWeight: 600,
  fontSize: 'larger',
  textAlign: 'center',

};

export function createItemDetail( item: IItemDetail, siteUrl: string, onClick?: any ) {

  let rows = [];
  
  ['versionlabel','sizeLabel','created','author','modified','editor', 'checkedOutId','uniquePerms'].map( thisKey => {
    rows.push( createRowFromItem( item, thisKey ) );
  });

  ['MediaServiceLocation','MediaServiceOCR','MediaServiceAutoTags','MediaServiceKeyPoints','MediaLengthInSeconds'].map( thisKey => {
    rows.push( createRowFromItem( item, thisKey ) );
  });
  
  ['bucket','ContentTypeId','ContentTypeName','ServerRedirectedEmbedUrl','MediaLengthInSeconds', 'isFolder'].map( thisKey => {
    rows.push( createRowFromItem( item, thisKey ) );
  });

  let previewUrl = siteUrl + "/_layouts/15/getpreview.ashx?resolution=0&clientMode=modernWebPart&path=" +
    window.origin + item.FileRef + "&width=500&height=400";

  let table = <div style={{marginRight: '10px'}} onClick={ onClick }>
      <h3 style={{ textAlign: 'center' }}>{ <Icon iconName= { item.iconName } style={ { fontSize: 'larger', color: item.iconColor, padding: '0px 15px 0px 0px', } }></Icon> }
        { item.FileLeafRef }</h3>
      {/* <table style={{padding: '0 20px'}}> */}
    <table style={{ tableLayout:"fixed" }} id="Select-b">
    { rows }

  </table>
  <div style = {{ paddingTop: '30px', paddingLeft: '20px' }}>
      <img src={ previewUrl } alt=""/>
    </div>
</div>;
return table;

}

function createRowFromItem( item: IItemDetail, key: string, format?: string, ) {
  let textValue = null;
  switch (key) {
    case 'author':
      textValue = `(${item.authorId}) ${item.authorTitle}`;
      break;
  
    case 'editor':
      textValue = `(${item.editorId}) ${item.editorTitle}`;
      break;
  
    case 'created':
      textValue = `${item.created.toLocaleString()}`;
      break;
  
    case 'modified':
      textValue = `${item.modified.toLocaleString()}`;
      break;
    
    case 'id':
      textValue = `Id: ${ item.id } Batch Details: ${ item.batch } ${ item.index }`;
      break;

    default:
      textValue = item[ key ];
      break;
  }

  if ( textValue ) {
    return <tr><td style={cellMaxStyle}>{ key }</td><td>{ textValue }</td></tr>;
  } else {
    return null;
  }
  
}
