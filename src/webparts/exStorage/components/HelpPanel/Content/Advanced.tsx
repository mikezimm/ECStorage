import * as React from 'react';

import styles from '../Component/InfoPane.module.scss';

//import { IHelpTableRow, IHelpTable, IPageContent, ISinglePageProps } from '../Component/ISinglePageProps';
import { IHelpTableRow, IHelpTable, IPageContent, ISinglePageProps } from '@mikezimm/npmfunctions/dist/HelpInfo/Component/ISinglePageProps';

export function advancedContent() {


    let html1 = <div>
        <h3>Sharing and Permission Test notes</h3>
        <ul>
            <li>If you share something, then remove sharing, the SharedWith info still remains</li>
            <li>If you copy an item that was shared, the copy is not shared, but the SharedWith properties still show on those items.</li>

            <li>If you Share, the item permission are automatically broken.</li>
            <li>However, if you find the 'Stop Sharing' button in Manage access, it removes the specific link permissions but the item still has broken permissions.</li>


            {/* <li></li>
            <li></li>
            <li></li>
            <li></li>
            <li></li>
            <li></li> */}
        </ul>
    </div>;

    return { html1: html1 };

}
  

