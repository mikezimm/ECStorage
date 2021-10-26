import * as React from 'react';

//import { IHelpTableRow, IHelpTable, IPageContent, ISinglePageProps } from '../Component/ISinglePageProps';
import { IHelpTableRow, IHelpTable, IPageContent, ISinglePageProps } from '@mikezimm/npmfunctions/dist/HelpInfo/Component/ISinglePageProps';

export const panelVersionNumber = '2021-10-27 -  1.0.0.20'; //Added to show in panel

export function aboutTable() {

    let table : IHelpTable  = {
        heading: 'Version History',
        headers: ['Date','Version','Focus'],
        rows: [],
    };

    table.rows.push( createAboutRow('2021-10-27', '1.0.0.20', `Standardize Banner code and options` ) );
    table.rows.push( createAboutRow('2021-10-22', '1.0.0.19', `Add CheckedOut and IsMinor to Versions tab` ) );
    table.rows.push( createAboutRow('2021-10-18', '1.0.0.18', `Add Version Info - Files in draft, high version counts` ) );
    table.rows.push( createAboutRow('2021-10-15', '1.0.0.17', `Misc data and styling improvements` ) );
    table.rows.push( createAboutRow('2021-10-13', '1.0.0.16', `Add Shared Events,Folder & Permission Details, improve items pages with click filtering,` ) );
    table.rows.push( createAboutRow('2021-10-04', '1.0.0.15', `Add Timeline tab (grid charts), Items Date flag style, Labels, styling` ) );
    table.rows.push( createAboutRow('2021-09-31', '1.0.0.14', `npmFunctions update.` ) );
    table.rows.push( createAboutRow('2021-09-30', '1.0.0.12', `Fix typos in help, Add Tricks logic, improve analytics` ) );
    table.rows.push( createAboutRow('2021-09-29', '1.0.0.11', `Update Banner to show actual webpart specific help, add Analytics` ) );
    table.rows.push( createAboutRow('2021-09-28', '1.0.0.10', `Add Library Dropdown, Improve Summary, Dups, Items, Preview, Media tags` ) );

    /*
    table.rows.push( ['2021-00-00', '1.0.0.0',    <span>Add support to view <b>List attachments, List link, Stat chart updates</b></span>,    ''] );
    */
    
    return { table: table };

}

function createAboutRow( date: string, version: string, focus: any ) {
    let tds = [<span style={{whiteSpace: 'nowrap'}} >{ date }</span>, 
        <span style={{whiteSpace: 'nowrap'}} >{ version }</span>, 
        <span>{ focus }</span>,] ;

    return tds;
}