import * as React from 'react';

import styles from '../banner/SinglePage/InfoPane.module.scss';

//import { IHelpTableRow, IHelpTable, IPageContent, ISinglePageProps } from '../Component/ISinglePageProps';
import { IHelpTableRow, IHelpTable, IPageContent, ISinglePageProps } from '../banner/SinglePage/ISinglePageProps';

export function basicsContent() {

    let html1 = <div>
        <h3>Summary</h3>
        <ul>
            <li>Highlights file metrics for a topic (like old or large files).</li>
            <li>Found in numerious tabs including: Main webpart, Users, Size, Age You, Dups, Folders.</li>
        </ul>

        <h3>Types</h3>
        <ul>
            <li>See what file types are in your library, ordered by quantity and size.</li>
            <li>Search by file extension, click on types to see largest files for a file type.</li>
            <li>Top of tab highlights some key metrics to know what file types are impacting your storage most.</li>
        </ul>

        <h3>Users</h3>
        <ul>
            <li>See what users are generating the most content.  By Qty and total size of files.</li>
            <li>Search by name, click on user to see file details on a specific user.</li>
        </ul>

        <h3>You</h3>
        <ul>
            <li>Highlights the files you created.</li>
            <li>Has same tabs as main webpart except Users.</li>
            <li>Provides same details as if you clicked on a user in the Users tab.</li>
        </ul>

        <h3>Size</h3>
        <ul>
            <li>Highlights the largest files in your library.</li>
            <li>Includes tabs for Summary and several size categories.</li>
            <li>Indicates when large files were created or last modified.</li>
            <li>Size Categories have counter in the tab to tell you how many files are in a bucket.</li> 
        </ul>
        <ul style={{ listStyleType: 'none'}}>
            <h4>Click on the category to see those files sorted largest first.</h4>
            <li>You can search for files by name, created user name, created date.</li>
            <li>Click the maginfing glass icon to see file details including media tags and preview.</li>
            <li>Click on the folder icon to go directly to that file's folder in the browser.</li>
            <li>Click on the file name to show preview of the file or go to the file.  Note some files that don't have previews are also downloaded such as zip.</li>
        </ul>

        <h3>Perms</h3>
        <ul>
            <li>Higlights items with unique permissions.</li>
        </ul>

        <h3>Dups</h3>
        <ul>
            <li>Higlights duplicate files.  Duplicate files are defined as files with the same filename.extension in multiple places.</li>
            <li>Click on a 'Dup' file name to see all the versions of the file.</li>
            <li>In the duplicate versions list, you get the same ability to see file details, click into folder or preview the file.</li>   
        </ul>

        <h3>Folders</h3>
        <li>Higlights Folders in your library.</li>

    </div>;

    return { html1: html1 };

}
  

