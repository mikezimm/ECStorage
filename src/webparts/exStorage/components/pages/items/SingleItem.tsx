import * as React from 'react';
import styles from '../../ExStorage.module.scss';
import { IEsItemsProps } from './IEsItemsProps';
import { IEsItemsState } from './IEsItemsState';
import { escape } from '@microsoft/sp-lodash-subset';


import { sp, Views, IViews, ISite } from "@pnp/sp/presets/all";
import { Web, IList, Site } from "@pnp/sp/presets/all";

import { ProgressIndicator } from 'office-ui-fabric-react/lib/ProgressIndicator';

import { Icon  } from 'office-ui-fabric-react/lib/Icon';

import { IItemDetail, IDuplicateFile } from '../../IExStorageState';
import { getFocusableByIndexPath } from 'office-ui-fabric-react';
  
const cellMaxStyle: React.CSSProperties = {
  whiteSpace: 'nowrap',
  height: '15px',
  padding: '10px 30px 0px 0px',
  fontWeight: 600,
  fontSize: 'larger',
  textAlign: 'center',

};

export function createItemDetail( item: IItemDetail, siteUrl: string, textSearch: string, onClick?: any, onPreviewClick?: any ) {

  let rows = [];
  
  ['id','versionlabel','sizeLabel','created','author','modified','editor', 'checkedOutId','uniquePerms','parentFolder'].map( thisKey => {
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
      <h2 style={{  }}>{ <Icon iconName= { item.iconName } style={ { fontSize: 'larger', color: item.iconColor, padding: '0px 15px 0px 0px', } }></Icon> }
        { item.FileLeafRef }</h2>
      {/* <table style={{padding: '0 20px'}}> */}

    <table style={{ tableLayout:"fixed" }} id="Select-b">
      { rows }
    </table>
      
    <div style = {{ paddingTop: '40px', display: 'flex', alignItems: 'flex-start', flexDirection: 'row' }}>
      <div>
        <div style={{ fontSize: 'larger', fontWeight: 600, paddingBottom: '20px'  }}>Preview (if available)"</div>
        <img src={ previewUrl } alt=""/>
      </div>

      {
        !textSearch || textSearch.length === 0 ? null :
        <div style = {{ paddingLeft: '50px', }}>
          <div style={{ fontSize: 'larger', fontWeight: 600  }}>Found by Searching for:</div>
          <p> { textSearch } </p>

          <div style={{ fontSize: 'larger', fontWeight: 600  }}>In this:</div>
          <div>
            <p>{ getHighlightedText( getItemSearchString( item ), textSearch ) }</p>
          </div>
        </div>
      }
    </div>

  </div>;
  return table;

}

export function createDuplicateDetail( item: IDuplicateFile, siteUrl: string, textSearch: string, onClick?: any, onPreviewClick?: any ) {

  let rows = [];
  
  [ 'sizeLabel','created','author','modified','editor','uniquePerms'].map( thisKey => {
    rows.push( createRowFromDup( item, thisKey ) );
  });
  
  ['ContentTypeId','ContentTypeName','ServerRedirectedEmbedUrl', 'isFolder'].map( thisKey => {
    rows.push( createRowFromDup( item, thisKey ) );
  });

  let table = <div style={{marginRight: '10px'}} onClick={ onClick }>
      <h2 style={{  }}>{ <Icon iconName= { item.iconName } style={ { fontSize: 'larger', color: item.iconColor, padding: '0px 15px 0px 0px', } }></Icon> }
        { item.FileLeafRef }</h2>

    <table style={{ tableLayout:"fixed" }} id="Select-b">
      { rows }
    </table>

    <div style = {{ paddingTop: '40px', display: 'flex', alignItems: 'flex-start', flexDirection: 'row' }}>
      {
        !textSearch || textSearch.length === 0 ? null :
        <div style = {{ paddingLeft: '50px', }}>
          <div style={{ fontSize: 'larger', fontWeight: 600  }}>Found by Searching for:</div>
          <p> { textSearch } </p>

          <div style={{ fontSize: 'larger', fontWeight: 600  }}>In this:</div>
          <div>
            <p>{ getHighlightedText( item.FileLeafRef , textSearch ) }</p>
          </div>
        </div>
      }
    </div>

  </div>;
  return table;

}

/**
 * Super cool solution based on:  https://stackoverflow.com/a/43235785
 * @param text 
 * @param highlight 
 */
export function getHighlightedText(text, highlight) {
  // Split on highlight term and include term into parts, ignore case
  const parts = text.split(new RegExp(`(${highlight})`, 'gi'));
  return <span> { parts.map((part, i) => 
      <span key={i} style={part.toLowerCase() === highlight.toLowerCase() ? { fontWeight: 'bold', backgroundColor: 'yellow' } : {} }>
          { part }
      </span>)
  } </span>;
}

export function getItemSearchString ( item: IItemDetail ) {

  let createdDate = new Date( item.created );
  let searchThis = [item.FileLeafRef, item.authorTitle, item.editorTitle, createdDate.toLocaleDateString() ].join('|');

  if ( item.MediaServiceAutoTags ) { searchThis += `|${item.MediaServiceAutoTags}` ; } //MSAT:
  if ( item.MediaServiceKeyPoints ) { searchThis += `|:${item.MediaServiceKeyPoints}` ; } //MSKP:
  if ( item.MediaServiceLocation ) { searchThis += `|:${item.MediaServiceLocation}` ; } //MSL:
  if ( item.MediaServiceOCR ) { searchThis += `|:${item.MediaServiceOCR}` ; } //MSOCR:

  return searchThis;

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
    return <tr><td style={cellMaxStyle}>{ key }</td><td style={{ padding: '10px 30px 0px 10px', }}>{ textValue }</td></tr>;
  } else {
    return null;
  }
  
}


function createRowFromDup( item: IDuplicateFile, key: string, format?: string, ) {
  let textValue = null;
  textValue = item[ key ];

  if ( textValue ) {
    return <tr><td style={cellMaxStyle}>{ key }</td><td style={{ padding: '10px 30px 0px 10px', }}>{ textValue }</td></tr>;
  } else {
    return null;
  }
  
}
