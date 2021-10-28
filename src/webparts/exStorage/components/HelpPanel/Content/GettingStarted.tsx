import * as React from 'react';

import styles from '../banner/SinglePage/InfoPane.module.scss';

//import { IHelpTableRow, IHelpTable, IPageContent, ISinglePageProps } from '../Component/ISinglePageProps';
import { IHelpTableRow, IHelpTable, IPageContent, ISinglePageProps } from '../banner/SinglePage/ISinglePageProps';

export function gettingStartedContent() {

    let html1 = <div>

        <h3>Add Webpart to your site</h3>
        <ol>
            <li>Go to Site Contents</li>
            <li>Click New, Add an App</li>
            <li>On left, click From Your Organization</li>
            <li>Search for Extreme Storage</li>
            <li>Click Add to Site</li>
        </ol>

        <h3>Add Webpart to your page</h3>
        <ol>
            <li>Go to a site page or create a new one</li>
            <li>Click Edit (page) in the upper right of page</li>
            <li>Insert new full width section (click + in left side of page)</li>
            <li>Insert new webpart (click + in center of section)</li>
            <li>Search for Extreme Storage</li>
            <li>Select webpart to add it</li>
        </ol>

        <h3>What to do?</h3>
        <ul>
            <li>If the library has less than 5k items, it auto-loads.</li>
            <li>If not, slide 'Fetch up to x files' slider to the right, then press 'Begin' button.</li>
            <li>When it's done, click around and explore, you can't break anything.</li>
        </ul>

        <h3>Webpart Options (in webpart settings)</h3>
        <ul>
            <li>Webpart defaults to the current document library</li>
            <li>You can point it to any document library (type in title in setting)</li>
            <li>Toggle Show Lists Dropdown in webpart - to let user change libraries</li>
            <li>Toggle Show System Lists - to include those lists in dropdown</li>
            <li>Exclude these from dropdown - Type in lists to exclude, use semi-colon separated words from title</li>
            <li>Include Media Tags in Search - Includes Media Tags including Location, OCR or other tags added in SharePoint</li>
        </ul>

        <h3>Other Options (in webpart settings)</h3>
        <p className="">All other webpart settings are advanced.  Please contact your SharePoint team for assistance if needed.</p>
    </div>;

    return { html1: html1 };

}
  

