import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { PropertyFieldOrder } from '@pnp/spfx-property-controls/lib/PropertyFieldOrder';
import {
  BaseClientSideWebPart,
  
  PropertyPaneTextField,
  PropertyPaneSlider,
  
  PropertyPaneDropdownOptionType
} from '@microsoft/sp-webpart-base';

import {
  PropertyPaneDropdown,
  IPropertyPaneConfiguration,
  PropertyPaneChoiceGroup,
  PropertyPaneButton,
  PropertyPaneToggle,
  IPropertyPaneDropdownOption,
  PropertyPaneHorizontalRule
} from '@microsoft/sp-property-pane';
import { FilePicker, IFilePickerResult , IFilePickerProps } 
from '@pnp/spfx-controls-react/lib/FilePicker';

import * as strings from 'IESModernSiteListFlipFlopWebPartStrings';
import IESModernSiteListFlipFlop from './components/IESModernSiteListFlipFlop';
import { IIESModernSiteListFlipFlopProps, IPropertyControlsTestWebPartProps } from './components/IESModernSiteListFlipFlopProps';
import ConfigureWebPart from './components/ConfigureWebPart/ConfigureWebPart';
import { sp } from '@pnp/sp';
// import "@pnp/sp/webs";
import { ListService } from '../../Services/ListService';
import { IImageStyleProps } from 'office-ui-fabric-react/lib/components/Image/Image.types';
import { spLists, spList } from './components/spLists';
import styles from './components/IESModernSiteListFlipFlop.module.scss';
import { css } from 'office-ui-fabric-react/lib/Utilities';
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';

export interface IImageGalleryWebPartProps {
  imageLibrary: string;
  // pageSize: number;
  style : string;
  theme : string;
  fontstyle: string;
  toBeFlip: boolean;
  FlipCardAxis: string;
  collectionData: any[]; 
  fontsize: string;
  fontposition: string;
   galleryAlignment: string;
  cardOrder: any[];

}
export interface IPnpPropertyFieldCollectionDataProps {  
  description: string;  
  collectionData: any[];  
}

export enum PropertyPaneButtonType{
  Normal,
  Primary,
  Hero,
  Compound,
  Command,
  Icon
}


export default class ImageGalleryWebPart extends BaseClientSideWebPart<IImageGalleryWebPartProps> {
  private listService: ListService
  private lists: IPropertyPaneDropdownOption[];
  private listsDropdownDisabled: boolean = true;
  private listDropDownOptions: IPropertyPaneDropdownOption[] =[]; 
  private listThemeColorDropDownOptions: IPropertyPaneDropdownOption[] =[];
  private listThemeColorChoiceOptions: any[] =[];
  private fontSizeDropDownOptions: IPropertyPaneDropdownOption[] =[]; 
  private ThemeWidth = 50;
  private SecondaryTeal = 'Secondary Teal';
  private SecondaryBlue = 'Secondary Blue';
  private SecondaryRed = 'Secondary Red';
  private SecondaryBlueGrey = 'Secondary  Blue Grey';
  private CDNBaseUrl = '';

  protected async onInit(): Promise<void> {
    const _ = await super.onInit();

    this.listService = new ListService(this.context.spHttpClient);

    sp.setup({
      spfxContext: this.context
    });

    sp.web.get().then(x => {
      console.log(x);
    });


    let cdnBaseUrl =
      this.context.pageContext.legacyPageContext.publicCdnBaseUrl;
    this.CDNBaseUrl = cdnBaseUrl;

    let edgeBrowser: boolean = window.navigator.userAgent.indexOf('Edg') != -1;
    let chromeBrowser: boolean = window.navigator.userAgent.indexOf('Chrome') != -1;


    if (edgeBrowser == false && chromeBrowser == false) {
      this.SecondaryTeal = '   Secondary Teal';
      this.SecondaryBlue = '  Secondary Blue';
      this.SecondaryBlueGrey = '  Secondary  Blue Grey';
      this.SecondaryRed = '   Secondary Red';

    }


    // let listresturl: string = this.context.pageContext.web.absoluteUrl + "/_api/web/lists?$select=Id,Title";
    // this.listService.getListsInfo(listresturl)
    //   .then((lists: spList[]) => {
    //     lists.forEach((list: spList) => {
    //       //Loads the drop down values  
    //       this.listDropDownOptions.push({
    //         key: list.Title, text: list.Title
    //       });
    //     });
    //   });

    // let listThemeColorurl: string = this.context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('iConnectThemeColor')/Items?$select=Id,Title,Color,DarkColor,LightColor";
    // this.listService.getListsInfo(listThemeColorurl)
    //   .then((lists: spList[]) => {
    //     lists.forEach((list: spList) => {
    //       //Loads the drop down values  

    //       this.listThemeColorChoiceOptions.push({ key: list.Title, text: list.Title });
    //     });
    //   });


  }

 

  public render(): void {
    
    sp.setup({
      spfxContext: this.context
    });

    

    let element: any;

    if (this.properties.collectionData) {

      element = React.createElement<IIESModernSiteListFlipFlopProps>(
        IESModernSiteListFlipFlop,
        {
          listName: this.properties.imageLibrary,
          context: this.context,
          siteUrl: this.context.pageContext.site.absoluteUrl,
          
          // pageSize: this.properties.pageSize,
          sizeName: this.properties.style,
          themecolorName: this.properties.theme,
          collectionData: this.properties.collectionData,
          enableFlip: this.properties.toBeFlip,   
          flipCardAxis: this.properties.FlipCardAxis,
          fontsize: this.properties.fontsize,
          fontposition: this.properties.fontposition,
          galleryAlignment: this.properties.galleryAlignment || 'left'

        }
      );
    }
    else {
      // show configure web part react component
      element = React.createElement(
        ConfigureWebPart,
        {
          webPartContext: this.context,
          title: "Flip Card",
          description: strings.MissingListConfiguration,
          buttonText: strings.ConfigureWebpartButtonText
        }
      );
    }

    //this.GetLists(); 
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  private validateTitle(value: string): string  {

    // if (value === null ||
    //   value.trim().length === 0) {
    //   return 'Kindly provide a Title';
    // }

    // if (value.length > 40) {
    //   return 'Title should not be longer than 40 characters';
    // }

    return '';
  }

  private validateBackText(value: string): string  {

    // if (value === null ||
    //   value.trim().length === 0) {
    //   return 'Kindly provide a Backtext';
    // }

    // if (value.length > 100) {
    //   return 'Backtext should not be longer than 100 characters';
    // }

    return '';
  }

  public CreateHref(filevalue: any){
    onCustomRender: (field: { id: any; }, _value: any, onUpdate: (arg0: any, arg1: string) => void, item: any, itemId: any, onError: (arg0: any, arg1: string) => void) => {
      return (
        React.createElement("div", null,
          React.createElement("a", { key: itemId, value: filevalue.fileAbsoluteUrl, onChange: (event: React.FormEvent<HTMLInputElement>) => {
            if (event.currentTarget.value === "error") {
              onError(field.id, "Value shouldn't be equal to error");
            } else {
              onError(field.id, "");
            }
            onUpdate(field.id, event.currentTarget.value);
          }}),
        )
      );
    }
    return "";
  }

  public saveIntoSharePoint(file: IFilePickerResult): Promise<any> {
    if (file.fileAbsoluteUrl == null) {

      return new Promise((resolve) => {

        file.downloadFileContent()
        .then(r => {
          sp.setup({
            spfxContext: this.context
          });
    
          sp.web.get().then( x => {
            console.log(x);
            let fileresult =  sp.web.getFolderByServerRelativeUrl(x.ServerRelativeUrl + "/SiteAssets").files.add(file.fileName, r, true);         
           
            resolve(fileresult);

          });
        });


      });
      file.downloadFileContent()
        .then(r => {
          sp.setup({
            spfxContext: this.context
          });
    
          sp.web.get().then( x => {
            console.log(x);
            let fileresult =  sp.web.getFolderByServerRelativeUrl(x.ServerRelativeUrl + "/SiteAssets").files.add(file.fileName, r, true);
            // this.setState({ ImageURL: document.location.origin + fileresult.data.ServerRelativeUrl });
           
            return fileresult;

            // fileresult.then(async result=>{

            //       return (result);
            //   // if (event.currentTarget.value === "error") {
            //   //   onError(field.id, "Value shouldn't be equal to error");
            //   // } else {
            //   //   onError(field.id, "");
            //   // }


            // });
          });


         
        });
    }
    else {
   
     return null;
    }



  }

  
  
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: "Flip Card Settings 3.0"
          },
          // ADDED: to turn groups into accordions
          displayGroupsAsAccordion: true,

          groups: [
            {
              groupName: 'Manage Links',
              isCollapsed: true,
              groupFields: [
             
              PropertyFieldCollectionData("collectionData", {
                key: "collectionData",
                label: "",
                panelHeader: "Flip Card Configuration",
                manageBtnLabel: "Manage Links",
                value: this.properties.collectionData,
                fields: [
                  {
                    id: "Title",
                    title: "Title",
                    type: CustomCollectionFieldType.string,
                    onGetErrorMessage: this.validateTitle.bind(this),
                    onCustomRender: (field, value, onUpdate, item, itemId, onError) => {
                      return (
                        React.createElement("div", null,
                          React.createElement("input", { key: itemId, value: value, onChange: (event: React.FormEvent<HTMLInputElement>) => {
                            if (event.currentTarget.value === "error") {
                              onError(field.id, "Value shouldn't be equal to error");
                            } else {
                              onError(field.id, "");
                            }
                            onUpdate(field.id, event.currentTarget.value);
                          }}),
                        )
                      );
                    },
                    required: true
                  },
                  {
                      id: "filePicker",
                      title: "Select Image",
                      type: CustomCollectionFieldType.custom,
                      required: true,
                      onCustomRender: (field, value, onUpdate, item, itemId, onError) => {
                        return (                         
                          React.createElement("span", null, "",
                          React.createElement("a", value!= null ? (value.fileAbsoluteUrl !=null ? {href: value.fileAbsoluteUrl} : {href: value.ServerRelativeUrl}) : "",
                          (value!= null ? (value.fileName != null ?   value.fileName :  value.Name): "")),                          
                          React.createElement("span", null, "",
                          React.createElement(FilePicker, {
                            key: itemId,                           
                            context: this.context,
                            buttonLabel:"Select File", 
                            hideLinkUploadTab: false,
                            storeLastActiveTab: true,
                            bingAPIKey: "<BING API KEY>",
                            accepts: [".gif", ".jpg", ".jpeg", ".bmp", ".dib", ".tif", ".tiff", ".ico", ".png", ".jxr", ".svg"],
                            buttonIcon:"FileImage",                                                    
                            onChanged: (selectedFile: IFilePickerResult) => {
                              onUpdate(field.id, selectedFile);                             
                              return event;
                            },
                            onSave: (filePickerResult: IFilePickerResult) => {

                                  //if fileAbsoluteUrl is null then its from Local system path
                              if(filePickerResult.fileAbsoluteUrl == null){
                                 // write code to save file to SharePoint doc library
                                //Upload new files from local system to sharepoint and get fileresult
                                this.saveIntoSharePoint(filePickerResult).then(fileResult=>{
                                  onUpdate(field.id, fileResult.data);
                                  return event; 
                                });                     

                              }else{//if fileAbsoluteUrl is not null then its from SharePoint Site Path

                                onUpdate(field.id, filePickerResult);
                                return event;
                              }                             
                            
                             
                            }
                          }
                          ))))
                      }
                   },                 
                  {
                    id: "RedirectLink",
                    title: "RedirectLink",
                    type: CustomCollectionFieldType.string,
                    required: true,
                    onCustomRender: (field, value, onUpdate, item, itemId, onError) => {
                      return (
                        React.createElement("div", null,
                          React.createElement("input", { key: itemId, value: value, onChange: (event: React.FormEvent<HTMLInputElement>) => {
                            if (event.currentTarget.value === "error") {
                              onError(field.id, "Value shouldn't be equal to error");
                            } else {
                              onError(field.id, "");
                            }
                            onUpdate(field.id, event.currentTarget.value);
                          }}),
                        )
                      );
                    }
                  },
                  {
                    id: "BackText",
                    title: "BackText",
                    type: CustomCollectionFieldType.string,
                    required: true,
                    onGetErrorMessage: this.validateBackText.bind(this),
                    onCustomRender: (field, value, onUpdate, item, itemId, onError) => {
                      return (
                        React.createElement("div", null,
                          React.createElement("input", { key: itemId, value: value, onChange: (event: React.FormEvent<HTMLInputElement>) => {
                            if (event.currentTarget.value === "error") {
                              onError(field.id, "Value shouldn't be equal to error");
                            } else {
                              onError(field.id, "");
                            }
                            onUpdate(field.id, event.currentTarget.value);
                          }}),
                        )
                      );
                    }
                  }
                ],
                disabled: false
              }),
                              PropertyFieldOrder("collectionData", {
                  key: "collectionDataOrder",
                  label: "Reorder Flip Cards",
                  items: this.properties.collectionData || [],
                  textProperty: "Title",
                  properties: this.properties,
                  onPropertyChange: this.onPropertyPaneFieldChanged
                }),
              PropertyPaneToggle('toBeFlip', {
                key: 'toBeFlip',
                label: 'Enable Flip',
                onText: 'Yes',
                offText: 'No',
                checked: true
              }),    
              PropertyPaneChoiceGroup('FlipCardAxis', {
                label: 'Select Card Axis',
                options: [{
                  key: 'X',
                  text: 'X'
                                 
                },
                {
                  key: 'Y',
                  text: 'Y',
                  checked: true               
                
                }
              ]       
              })            
              ]
            },
            {
              groupName: "Manage Appearance",
              // ADDED: to collapse group initially
              isCollapsed: true,
              // END added
              groupFields: [
                PropertyPaneChoiceGroup('galleryAlignment', {
                  label: 'Select Gallery Alignment',
                  options: [{
                    key: 'left',
                    text: 'left'
                  },
                  {
                    key: 'center',
                    text: 'Center'
                  },
                  {
                    key: 'right',
                    text: 'right'

                  }
                    // ,
                    // {
                    //   key: 'center',
                    //   text: 'center'
                    // }
                  ]
                }),
                PropertyPaneChoiceGroup('style', {
                  label: 'Select Card Size',
                  options: [{
                    key: 'Width:212px;Height:156px',
                    text: 'Compact'
                    // checked: true
                  },
                  {
                    key: 'Width:256px;Height:166px',
                    text: 'Medium'
                  
                  },
                  {
                    key: 'Width:314px;Height:226px',
                    text: 'Large'
                  },
                  {
                    key: 'Width:195px;Height:75px',
                    text: 'Button'
                  }
                  
                ]       
                }),
                PropertyPaneChoiceGroup('theme', {
                  label: 'Select Theme Color',
                  options: [{
                    key: 'Dark:#e71316;Light:#f16c6e',
                    text: 'ThermoFisher Theme',
                    ariaLabel: 'ariaLabel',
                    // imageSrc: 'https://thermofisher.sharepoint.com/sites/ModernApps/Images1/ThemeColor/32/ThermoFisherTheme.jpg',
                    imageSrc: `${this.CDNBaseUrl}/thermofisher.sharepoint.com/CDN/publishingimages/ModernThemes/TFSTheme.jpg`,
                    imageSize: { width: this.ThemeWidth, height: 32 },
                    // selectedImageSrc: 'https://publiccdn.sharepointonline.com/thermofisher.sharepoint.com/CDN/publishingimages/ModernThemes/TFSTheme.jpg',
                    checked: true                                     
                  },
                  {
                    key: 'Dark:#f12e2c;Light:#f7807e',
                    text: 'Primary Red',
                    imageSrc: `${this.CDNBaseUrl}/thermofisher.sharepoint.com/CDN/publishingimages/ModernThemes/TFSTheme.jpg`,
                    // imageSrc: 'https://publiccdn.sharepointonline.com/thermofisher.sharepoint.com/CDN/publishingimages/ModernThemes/PrimaryRed.jpg',
                    imageSize: { width: this.ThemeWidth, height: 32 },
                    // selectedImageSrc: 'https://publiccdn.sharepointonline.com/thermofisher.sharepoint.com/CDN/publishingimages/ModernThemes/PrimaryRed.jpg'
                  },
                  {
                    key: 'Dark:#005bac;Light:#5293cd',
                    text: 'Primary Blue',
                    imageSrc: `${this.CDNBaseUrl}/thermofisher.sharepoint.com/CDN/publishingimages/ModernThemes/PrimaryBlue.jpg`,
                    // imageSrc: PrimaryBlue,
                    imageSize: { width: this.ThemeWidth, height: 32 },
                    // selectedImageSrc: 'https://publiccdn.sharepointonline.com/thermofisher.sharepoint.com/CDN/publishingimages/ModernThemes/PrimaryBlue.jpg'
                  },
                  {
                    key: 'Dark:#595959;Light:#9c9c9c',
                    text: 'Primary Grey',
                    imageSrc: `${this.CDNBaseUrl}/thermofisher.sharepoint.com/CDN/publishingimages/ModernThemes/PrimaryGrey.jpg`,
                    imageSize: { width: this.ThemeWidth, height: 32 },
                    // selectedImageSrc: 'https://publiccdn.sharepointonline.com/thermofisher.sharepoint.com/CDN/publishingimages/ModernThemes/PrimaryGrey.jpg'
                  },
                  {
                    key: 'Dark:#261f63;Light:#655fa2',
                    text: 'Primary Purple',
                    imageSrc: `${this.CDNBaseUrl}/thermofisher.sharepoint.com/CDN/publishingimages/ModernThemes/PrimaryPurple.jpg`,
                    imageSize: { width: this.ThemeWidth, height: 32 },
                    // selectedImageSrc: 'https://publiccdn.sharepointonline.com/thermofisher.sharepoint.com/CDN/publishingimages/ModernThemes/PrimaryPurple.jpg'
                  },
                  {
                    key: 'Dark:#7473c3;Light:#a5a5da',
                    text: 'Secondary Purple',
                    imageSrc: `${this.CDNBaseUrl}/thermofisher.sharepoint.com/CDN/publishingimages/ModernThemes/SecondaryPurple.jpg`,
                    imageSize: { width: this.ThemeWidth, height: 32 },
                    // selectedImageSrc: 'https://publiccdn.sharepointonline.com/thermofisher.sharepoint.com/CDN/publishingimages/ModernThemes/SecondaryPurple.jpg'
                  },
                  {
                    key: 'Dark:#016682;Light:#499db4',
                    text: this.SecondaryTeal,
                    imageSrc: `${this.CDNBaseUrl}/thermofisher.sharepoint.com/CDN/publishingimages/ModernThemes/SecondaryTeal.jpg`,
                    imageSize: { width: this.ThemeWidth, height: 32 },
                    // selectedImageSrc: 'https://publiccdn.sharepointonline.com/thermofisher.sharepoint.com/CDN/publishingimages/ModernThemes/SecondaryTeal.jpg'
                  },
                  {
                    key: 'Dark:#0e89e9;Light:#69b6f1',
                    text: this.SecondaryBlue,
                    imageSrc: `${this.CDNBaseUrl}/thermofisher.sharepoint.com/CDN/publishingimages/ModernThemes/SecondaryBlue.jpg`,
                    imageSize: { width: this.ThemeWidth, height: 32 },
                    // selectedImageSrc: 'https://publiccdn.sharepointonline.com/thermofisher.sharepoint.com/CDN/publishingimages/ModernThemes/SecondaryBlue.jpg'
                  },
                  {
                    key: 'Dark:#ad1e2d;Light:#ce6772',
                    text: this.SecondaryRed,
                    imageSrc: `${this.CDNBaseUrl}/thermofisher.sharepoint.com/CDN/publishingimages/ModernThemes/SecondaryRed.jpg`,
                    imageSize: { width: this.ThemeWidth, height: 32 },
                    // selectedImageSrc: 'https://publiccdn.sharepointonline.com/thermofisher.sharepoint.com/CDN/publishingimages/ModernThemes/SecondaryRed.jpg'
                  },
                  {
                    key: 'Dark:#5989b7;Light:#93b5d4',
                    text:  this.SecondaryBlueGrey,
                    imageSrc: `${this.CDNBaseUrl}/thermofisher.sharepoint.com/CDN/publishingimages/ModernThemes/SecondaryBlueGrey.jpg`,
                    imageSize: { width: this.ThemeWidth, height: 32 },
                    // selectedImageSrc: 'https://publiccdn.sharepointonline.com/thermofisher.sharepoint.com/CDN/publishingimages/ModernThemes/SecondaryBlueGrey.jpg'
                  }
                ]       
                }),
                PropertyPaneChoiceGroup('fontsize', {
                  label: 'Select Font Size',
                  options:[{                  
                      key: 'Small',
                      text: 'Small'
                  },{
                    key: 'Medium',
                    text: 'Medium'
                    // checked: true,
                  },{
                    key: 'Large',
                    text: 'Large'
                  }
                
                ]
                }),
                // PropertyPaneButton('', {
                //   text: "Icon button ('AddFriend' icon)",
                //   buttonType: PropertyPaneButtonType.Icon,
                //   icon: 'FontIncrease',
                //   onClick: this.fontSizeiconIncreaseButtonClick
                //  }),
                //  PropertyPaneHorizontalRule(),
                //  PropertyPaneButton('', {
                //   text: "Icon button ('AddFriend' icon)",
                //   buttonType: PropertyPaneButtonType.Icon,
                //   icon: 'FontDecrease',
                //   onClick: this.fontSizeiconDecreaseButtonClick(this.properties.fontsize)
                //  }),     
                //  PropertyPaneHorizontalRule(),
                //  PropertyPaneButton('', {
                //   text: "Icon button ('AddFriend' icon)",
                //   buttonType: PropertyPaneButtonType.Icon,
                //   icon: 'Bold',
                //   onClick: this.fontSizeiconDecreaseButtonClick(this.properties.fontsize)
                //  }),
                //  PropertyPaneHorizontalRule(),
                //  PropertyPaneButton('', {
                //   text: "Icon button ('AddFriend' icon)",
                //   buttonType: PropertyPaneButtonType.Icon,
                //   icon: 'Italic',
                //   onClick: this.fontSizeiconDecreaseButtonClick(this.properties.fontsize)
                //  }),
                PropertyPaneChoiceGroup('fontposition', {
                  label: 'Select Font Position',
                  options: [
                  
                    {
                      key: 'AlignLeft',
                      text: 'Align Left',
                      // checked: true,
                      iconProps: {
                        officeFabricIconFontName: 'AlignLeft'
                      },
                      imageSize: { width: 20, height: 20 }
                    }                    ,
                    
                    {
                      key: 'AlignCenter',
                      text: 'Align  Center',
                      iconProps: {
                        officeFabricIconFontName: 'AlignCenter'
                      },
                      imageSize: { width: 50, height: 20 }
                    },
                    {
                      key: 'AlignRight',
                      text: 'Align Right',
                      iconProps: {
                        officeFabricIconFontName: 'AlignRight'
                      },
                      imageSize: { width: 20, height: 20 }
                    }
                    
                  ]              
                }), 
              ]

            },
           
          ]
        }
      ]
    };
  }
}
