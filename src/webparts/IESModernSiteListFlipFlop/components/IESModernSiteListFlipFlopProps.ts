export interface IIESModernSiteListFlipFlopProps {
  listName: string;
  context: any;
  siteUrl:string;
  // pageSize:number;
  sizeName: string;
  themecolorName: string;
  fontsize: string;
  collectionData: any[];
  enableFlip: boolean;
  flipCardAxis: string;
  fontposition: string;
  galleryAlignment: string;
}

export interface spListItems{  
  value: spListItem[];  
}  
export interface spListItem{  
  Title: string;  
  Overview: string;
  FlipFlopBackDescription: string;
  RedirectHyperLink: string;
  HyperLink : string;
  BackgroundImageURL: string;
  DarkColor: string;
  LightColor: string;
  Width: string;
  Height: string;
  ID: string;  
  Created: string;  
  Author: {  
    Title: string;  
  };  
}  
  
export interface spList{  
Title:string; 
Overview: string; 
FlipFlopBackDescription: string;
RedirectHyperLink: string;
HyperLink : string;
BackgroundImageURL: string;
id: string;  
}  
export interface spLists{  
  value: spList[];  
} 

export interface IPropertyControlsTestWebPartProps {
  collectionData: any[];
}

