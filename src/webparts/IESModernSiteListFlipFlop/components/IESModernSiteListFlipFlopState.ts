import { IImage } from "../../../Interfaces";

export interface IIESModernSiteListFlipFlopState {
    showPanel: boolean;
    selectedImage?: IImage;
    showLoader: boolean;
    itemsNotFoundMessage?: string,
    sQuery?: string,
    dQuery?: string
    itemsNotFound?: boolean,
    itemCount?: number;
    // pageSize?: number;
    currentPage?: number;
    items?: any[];
    status?: string;
    nextLink: string;
  }

  export interface IPnpDropDownState {  
    RiskStatus: any[];  
  }  