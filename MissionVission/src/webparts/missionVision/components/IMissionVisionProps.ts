import { SPHttpClient } from '@microsoft/sp-http';

export interface IMissionVisionProps { 
  listName: string;
  siteUrl: string;
  spHttpClient: SPHttpClient;

  // Grid style
  alignment: string;
  gridBackground: string;
  gridGap: string;
  itemAsLink: boolean;

  // Item style
  itemWidth:string;
  itemBackground: string;
  iconWidth: string;
  titleFontSize: string;
  descFontSize: string;
}

export interface IGridItem {
  id: number;
  title: string;
  description: string;
  url: string;
  iconUrl: string;
}