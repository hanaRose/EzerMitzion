import * as React from 'react';
import styles from './SearchResultsTitle.module.scss';
import type { ISearchResultsTitleProps } from './ISearchResultsTitleProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class SearchResultsTitle extends React.Component<ISearchResultsTitleProps> {
  public render(): React.ReactElement<ISearchResultsTitleProps> {
    const folderName = new URLSearchParams(window.location.search).get("folderName") || "Shared Documents";
    return (
       <div>
          <div className={styles.wpTitle}>{folderName}</div>         
        </div>
    );
  }
}
