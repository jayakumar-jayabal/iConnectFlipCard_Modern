import * as React from 'react';
import styles from './IESModernSiteListFlipFlop.module.scss';
import { IIESModernSiteListFlipFlopProps } from './IESModernSiteListFlipFlopProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { css, classNamesFunction, IStyleFunction } from '@uifabric/utilities/lib';
import { RichText } from "@pnp/spfx-controls-react/lib/RichText";
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { IListService, IImage } from '../../../Interfaces';
import { ListService } from '../../../Services/ListService';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { objectDefinedNotNull, stringIsNullOrEmpty } from '@pnp/common';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { Button } from 'office-ui-fabric-react/lib/Button';
import { IIESModernSiteListFlipFlopState } from './IESModernSiteListFlipFlopState';


import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import { ScreenWidthMaxMedium, FontSizes } from 'office-ui-fabric-react/lib/Styling';


export default class ImageGallery extends React.Component<IIESModernSiteListFlipFlopProps, IIESModernSiteListFlipFlopState> {

  private _spService: IListService;
  private selectQuery: string[] = [];
  private expandQuery: string[] = [];
  private filterQuery: string[] = [];
  private urlCollection: string[] = [];
  private themeColor: string = "";
  public widthheight: string = "";
  private enableFlip: boolean;
  private flipCardAxis: string;
  private fontsizeStyle: string;
  private fontpositionStyle: string;
  private galleryAlignmentStyle: string;
  /**
 *
 */
  constructor(props: IIESModernSiteListFlipFlopProps, state: IIESModernSiteListFlipFlopState) {
    super(props);

    this.state = {
      items: [],
      showPanel: false,
      selectedImage: {} as IImage,
      showLoader: false,
      itemsNotFound: false,
      // pageSize: this.props.pageSize,
      currentPage: 1,
      nextLink: "",

    }


    this._spService = new ListService(this.props.context.spHttpClient);
  }


  public async componentDidMount() {
    //Get Images from the library 

    this.themeColor = this.props.themecolorName;
    this.widthheight = this.props.sizeName;
    this.enableFlip = this.props.enableFlip;
    this.flipCardAxis = this.props.flipCardAxis;
    this.fontsizeStyle = this.props.fontsize;
    this.fontpositionStyle = this.props.fontposition;
    this.galleryAlignmentStyle = this.props.galleryAlignment;

    //If themecolor is null or empty assign default theme
    if (this.props.themecolorName == "" || this.props.themecolorName == undefined) {

      this.themeColor = "Dark:#e71316;Light:#f16c6e";
    }

    if (this.props.sizeName == "" || this.props.sizeName == undefined) {

      this.widthheight = "Width:212px;Height:156px";
    }

    if (this.enableFlip == null || this.enableFlip == undefined) {

      this.enableFlip = true;

    }

    if (this.props.flipCardAxis == null || this.props.flipCardAxis == undefined) {

      this.flipCardAxis = "Y";
    }

    if (this.props.fontsize == null || this.props.fontsize == undefined) {

      this.fontsizeStyle = "Medium";
    }

    if (this.props.fontposition == null || this.props.fontposition == undefined) {

      this.fontpositionStyle = "AlignLeft";
    }

    if (this.props.galleryAlignment == null || this.props.galleryAlignment == undefined) {

      this.galleryAlignmentStyle = "left";
    }

  }

  private _onImageClick(selectedImage: any): void {

    window.open(selectedImage.RedirectLink);

  }

  public render(): React.ReactElement<IIESModernSiteListFlipFlopProps> {


    let result: React.ReactNode[] = [];

    let tagList;
    let flipWidthHeight;
    let figcaptionstyle;
    let backstyle;
    let frontstyle;
    let imageheight;
    let enableFlipFlop;
    let flipCardAxis;
    let fontsizeop: string;
    let fonttextalign: string;
    let CardSize: string;
    let galleryStyle: string;

    this.themeColor = this.props.themecolorName; //Dark:#005bac;Light:#5293cd
    this.widthheight = this.props.sizeName;
    this.enableFlip = this.props.enableFlip;
    this.flipCardAxis = this.props.flipCardAxis;
    this.fontsizeStyle = this.props.fontsize;
    this.fontpositionStyle = this.props.fontposition;
    this.galleryAlignmentStyle = this.props.galleryAlignment;

    //If themecolor is null or empty assign default theme
    if (this.props.themecolorName == "" || this.props.themecolorName == undefined) {

      this.themeColor = "Dark:#e71316;Light:#f16c6e";
    }

    if (this.props.sizeName == "" || this.props.sizeName == undefined) {

      this.widthheight = "Width:212px;Height:156px";
    }

    if (this.enableFlip == null || this.enableFlip == undefined) {

      this.enableFlip = true;
    }

    if (this.props.flipCardAxis == null || this.props.flipCardAxis == undefined) {

      this.flipCardAxis = 'Y';
    }

    if (this.props.fontsize == null || this.props.fontsize == undefined) {

      this.fontsizeStyle = 'Medium';
    }

    if (this.props.fontposition == null || this.props.fontposition == undefined) {

      this.fontpositionStyle = 'AlignLeft';
    }

     if (this.props.galleryAlignment == null || this.props.galleryAlignment == undefined) {

      this.galleryAlignmentStyle = 'left';
    }



    //Set Font Size
    if (this.fontsizeStyle == "Medium") {
      fontsizeop = "medium";
    } else if (this.fontsizeStyle == "Small") {
      fontsizeop = "small";
    } else if (this.fontsizeStyle == "Large") {
      fontsizeop = "large";
    } else {
      // if(this.props.fontsize!=null || this.props.fontsize != undefined){

      //   fontsizeop = this.props.fontsize;
      // }
      if (this.fontsizeStyle == "small") {
        fontsizeop = "x-small";
      }

    }

    //Set Font Text Alignment
    if (this.fontpositionStyle == "AlignJustify") {
      fonttextalign = "AlignJustify";
    } else if (this.fontpositionStyle == "AlignLeft") {
      fonttextalign = "AlignLeft";
    } else if (this.fontpositionStyle == "AlignRight") {
      fonttextalign = "AlignRight";
    } else if (this.fontpositionStyle == "AlignCenter") {
      fonttextalign = "AlignCenter";
    }


    //Set Gallery Text Alignment
    if (this.props.galleryAlignment == "left") {
      galleryStyle = "left";
    } else if (this.props.galleryAlignment == "right") {
      galleryStyle = "right";
    }
    else if (this.galleryAlignmentStyle == "center") {
      galleryStyle = "center";
    }


    const alignmentClass =
      galleryStyle === 'right'
        ? styles.alignRight
        : galleryStyle === 'center'
          ? styles.alignCenter
          : styles.alignLeft;

    //Set EnableFlipFlop
    enableFlipFlop = this.enableFlip;
    flipCardAxis = this.flipCardAxis;


    backstyle = {
      // color: 'red',
      background: 'linear-gradient(to bottom right, ' + this.themeColor.split(';')[0].toString().split(':')[1].toString() + ',  ' +
        this.themeColor.split(';')[1].toString().split(':')[1].toString() + ')',
    };

    frontstyle = {
      background: this.themeColor.split(';')[0].toString().split(':')[1].toString(),
      border: '0px solid ' + this.themeColor.split(';')[0].toString().split(':')[1].toString(),
    };



    flipWidthHeight = {
      //Width:212px;Height:156px
      width: this.widthheight.split(';')[0].toString().split(':')[1].toString(),
      height: this.widthheight.split(';')[1].toString().split(':')[1].toString(),
    }


    //Layout: Compact and Font Position

    if (this.widthheight.toString() == "Width:212px;Height:156px") {

      CardSize = "Compact";

    }

    if (this.widthheight.toString() == "Width:256px;Height:166px") {

      CardSize = "Medium";

    }


    if (this.widthheight.toString() == "Width:314px;Height:226px") {

      CardSize = "Large";

    }


    if (this.widthheight.toString() == "Width:195px;Height:75px") {

      CardSize = "ButtonSmall";

    }

    if (this.props.collectionData) {

      result = this.props.collectionData && this.props.collectionData.map((item, index) => {
        return (
          <div key={index} className={css(styles.column)} onClick={() => this._onImageClick(item)}>

            {
              enableFlipFlop == true && flipCardAxis == "Y" ?

                (

                  CardSize == "Compact" && fonttextalign == "AlignLeft" && fontsizeop == "small" ?
                    (
                      <div className={css(styles.flipCard)}>
                        <div className={css(styles.flipCardInner)}>
                          <div className={css(styles.flipCardFront)} style={frontstyle}>
                            <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                              className={css(styles.flipCardImageSizeCompact)} />
                            <div className={css(styles.flipCardTitleCompactSmall, styles.flipCardTitleCompactSmallLeft)}>{item.Title}</div>
                          </div>
                          <div className={css(styles.flipCardBack)} style={backstyle}>
                            <p className={css(styles.flipCardBackTextCompact, styles.flipCardTitleCompactSmallLeft)}>
                              {item.BackText}
                            </p>
                          </div>
                        </div>
                      </div>

                    ) : CardSize == "Compact" && fonttextalign == "AlignCenter" && fontsizeop == "small" ?
                      (
                        <div className={css(styles.flipCard)}>
                          <div className={css(styles.flipCardInner)}>
                            <div className={css(styles.flipCardFront)} style={frontstyle}>
                              <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                className={css(styles.flipCardImageSizeCompact)} />
                              <div className={css(styles.flipCardTitleCompactSmall, styles.flipCardTitleCompactSmallCenter)}>{item.Title}</div>
                            </div>
                            <div className={css(styles.flipCardBack)} style={backstyle}>
                              <p className={css(styles.flipCardBackTextCompact, styles.flipCardTitleCompactSmallCenter)}>
                                {item.BackText}
                              </p>
                            </div>
                          </div>
                        </div>

                      ) : CardSize == "Compact" && fonttextalign == "AlignRight" && fontsizeop == "small" ?
                        (
                          <div className={css(styles.flipCard)}>
                            <div className={css(styles.flipCardInner)}>
                              <div className={css(styles.flipCardFront)} style={frontstyle}>
                                <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                  className={css(styles.flipCardImageSizeCompact)} />
                                <div className={css(styles.flipCardTitleCompactSmall, styles.flipCardTitleCompactSmallRight)}>{item.Title}</div>
                              </div>
                              <div className={css(styles.flipCardBack)} style={backstyle}>
                                <p className={css(styles.flipCardBackTextCompact, styles.flipCardTitleCompactSmallRight)}>
                                  {item.BackText}
                                </p>
                              </div>
                            </div>
                          </div>

                        ) : CardSize == "Compact" && fonttextalign == "AlignLeft" && fontsizeop == "medium" ?
                          (
                            <div className={css(styles.flipCard)}>
                              <div className={css(styles.flipCardInner)}>
                                <div className={css(styles.flipCardFront)} style={frontstyle}>
                                  <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                    className={css(styles.flipCardImageSizeCompact)} />
                                  <div className={css(styles.flipCardTitleCompactMedium, styles.flipCardTitleCompactMediumLeft)}>{item.Title}</div>
                                </div>
                                <div className={css(styles.flipCardBack)} style={backstyle}>
                                  <p className={css(styles.flipCardBackTextCompact, styles.flipCardTitleCompactMediumLeft)}>
                                    {item.BackText}
                                  </p>
                                </div>
                              </div>
                            </div>

                          ) : CardSize == "Compact" && fonttextalign == "AlignCenter" && fontsizeop == "medium" ?
                            (
                              <div className={css(styles.flipCard)}>
                                <div className={css(styles.flipCardInner)}>
                                  <div className={css(styles.flipCardFront)} style={frontstyle}>
                                    <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                      className={css(styles.flipCardImageSizeCompact)} />
                                    <div className={css(styles.flipCardTitleCompactMedium, styles.flipCardTitleCompactMediumCenter)}>{item.Title}</div>
                                  </div>
                                  <div className={css(styles.flipCardBack)} style={backstyle}>
                                    <p className={css(styles.flipCardBackTextCompact, styles.flipCardTitleCompactMediumCenter)}>
                                      {item.BackText}
                                    </p>
                                  </div>
                                </div>
                              </div>

                            ) : CardSize == "Compact" && fonttextalign == "AlignRight" && fontsizeop == "medium" ?
                              (
                                <div className={css(styles.flipCard)}>
                                  <div className={css(styles.flipCardInner)}>
                                    <div className={css(styles.flipCardFront)} style={frontstyle}>
                                      <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                        className={css(styles.flipCardImageSizeCompact)} />
                                      <div className={css(styles.flipCardTitleCompactMedium, styles.flipCardTitleCompactMediumRight)}>{item.Title}</div>
                                    </div>
                                    <div className={css(styles.flipCardBack)} style={backstyle}>
                                      <p className={css(styles.flipCardBackTextCompact, styles.flipCardTitleCompactMediumRight)}>
                                        {item.BackText}
                                      </p>
                                    </div>
                                  </div>
                                </div>

                              ) : CardSize == "Compact" && fonttextalign == "AlignLeft" && fontsizeop == "large" ?
                                (
                                  <div className={css(styles.flipCard)}>
                                    <div className={css(styles.flipCardInner)}>
                                      <div className={css(styles.flipCardFront)} style={frontstyle}>
                                        <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                          className={css(styles.flipCardImageSizeCompact)} />
                                        <div className={css(styles.flipCardTitleCompactLarge, styles.flipCardTitleCompactLargeLeft)}>{item.Title}</div>
                                      </div>
                                      <div className={css(styles.flipCardBack)} style={backstyle}>
                                        <p className={css(styles.flipCardBackTextCompact, styles.flipCardTitleCompactLargeLeft)}>
                                          {item.BackText}
                                        </p>
                                      </div>
                                    </div>
                                  </div>

                                ) : CardSize == "Compact" && fonttextalign == "AlignCenter" && fontsizeop == "large" ?
                                  (
                                    <div className={css(styles.flipCard)}>
                                      <div className={css(styles.flipCardInner)}>
                                        <div className={css(styles.flipCardFront)} style={frontstyle}>
                                          <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                            className={css(styles.flipCardImageSizeCompact)} />
                                          <div className={css(styles.flipCardTitleCompactLarge, styles.flipCardTitleCompactLargeCenter)}>{item.Title}</div>
                                        </div>
                                        <div className={css(styles.flipCardBack)} style={backstyle}>
                                          <p className={css(styles.flipCardBackTextCompact, styles.flipCardTitleCompactLargeCenter)}>
                                            {item.BackText}
                                          </p>
                                        </div>
                                      </div>
                                    </div>

                                  ) : CardSize == "Compact" && fonttextalign == "AlignRight" && fontsizeop == "large" ?
                                    (
                                      <div className={css(styles.flipCard)}>
                                        <div className={css(styles.flipCardInner)}>
                                          <div className={css(styles.flipCardFront)} style={frontstyle}>
                                            <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                              className={css(styles.flipCardImageSizeCompact)} />
                                            <div className={css(styles.flipCardTitleCompactLarge, styles.flipCardTitleCompactLargeRight)}>{item.Title}</div>
                                          </div>
                                          <div className={css(styles.flipCardBack)} style={backstyle}>
                                            <p className={css(styles.flipCardBackTextCompact, styles.flipCardTitleCompactLargeRight)}>
                                              {item.BackText}
                                            </p>
                                          </div>
                                        </div>
                                      </div>

                                    ) : CardSize == "Medium" && fonttextalign == "AlignLeft" && fontsizeop == "small" ?
                                      (
                                        <div className={css(styles.flipCardMedium)}>
                                          <div className={css(styles.flipCardInner)}>
                                            <div className={css(styles.flipCardFront)} style={frontstyle}>
                                              <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                className={css(styles.flipCardImageSizeMedium)} />
                                              <div className={css(styles.flipCardTitleMediumSmall, styles.flipCardTitleMediumSmallLeft)}>{item.Title}</div>
                                            </div>
                                            <div className={css(styles.flipCardBack)} style={backstyle}>
                                              <p className={css(styles.flipCardBackTextMedium, styles.flipCardTitleMediumSmallLeft)}>
                                                {item.BackText}
                                              </p>
                                            </div>
                                          </div>
                                        </div>

                                      ) : CardSize == "Medium" && fonttextalign == "AlignCenter" && fontsizeop == "small" ?
                                        (
                                          <div className={css(styles.flipCardMedium)}>
                                            <div className={css(styles.flipCardInner)}>
                                              <div className={css(styles.flipCardFront)} style={frontstyle}>
                                                <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                  className={css(styles.flipCardImageSizeMedium)} />
                                                <div className={css(styles.flipCardTitleMediumSmall, styles.flipCardTitleMediumSmallCenter)}>{item.Title}</div>
                                              </div>
                                              <div className={css(styles.flipCardBack)} style={backstyle}>
                                                <p className={css(styles.flipCardBackTextMedium, styles.flipCardTitleMediumSmallCenter)}>
                                                  {item.BackText}
                                                </p>
                                              </div>
                                            </div>
                                          </div>

                                        ) : CardSize == "Medium" && fonttextalign == "AlignRight" && fontsizeop == "small" ?
                                          (
                                            <div className={css(styles.flipCardMedium)}>
                                              <div className={css(styles.flipCardInner)}>
                                                <div className={css(styles.flipCardFront)} style={frontstyle}>
                                                  <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                    className={css(styles.flipCardImageSizeMedium)} />
                                                  <div className={css(styles.flipCardTitleMediumSmall, styles.flipCardTitleMediumSmallRight)}>{item.Title}</div>
                                                </div>
                                                <div className={css(styles.flipCardBack)} style={backstyle}>
                                                  <p className={css(styles.flipCardBackTextMedium, styles.flipCardTitleMediumSmallRight)}>
                                                    {item.BackText}
                                                  </p>
                                                </div>
                                              </div>
                                            </div>

                                          ) : CardSize == "Medium" && fonttextalign == "AlignLeft" && fontsizeop == "medium" ?
                                            (
                                              <div className={css(styles.flipCardMedium)}>
                                                <div className={css(styles.flipCardInner)}>
                                                  <div className={css(styles.flipCardFront)} style={frontstyle}>
                                                    <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                      className={css(styles.flipCardImageSizeMedium)} />
                                                    <div className={css(styles.flipCardTitleMediumMedium, styles.flipCardTitleMediumMediumLeft)}>{item.Title}</div>
                                                  </div>
                                                  <div className={css(styles.flipCardBack)} style={backstyle}>
                                                    <p className={css(styles.flipCardBackTextMedium, styles.flipCardTitleMediumMediumLeft)}>
                                                      {item.BackText}
                                                    </p>
                                                  </div>
                                                </div>
                                              </div>

                                            ) : CardSize == "Medium" && fonttextalign == "AlignCenter" && fontsizeop == "medium" ?
                                              (
                                                <div className={css(styles.flipCardMedium)}>
                                                  <div className={css(styles.flipCardInner)}>
                                                    <div className={css(styles.flipCardFront)} style={frontstyle}>
                                                      <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                        className={css(styles.flipCardImageSizeMedium)} />
                                                      <div className={css(styles.flipCardTitleMediumMedium, styles.flipCardTitleMediumMediumCenter)}>{item.Title}</div>
                                                    </div>
                                                    <div className={css(styles.flipCardBack)} style={backstyle}>
                                                      <p className={css(styles.flipCardBackTextMedium, styles.flipCardTitleMediumMediumCenter)}>
                                                        {item.BackText}
                                                      </p>
                                                    </div>
                                                  </div>
                                                </div>

                                              ) : CardSize == "Medium" && fonttextalign == "AlignRight" && fontsizeop == "medium" ?
                                                (
                                                  <div className={css(styles.flipCardMedium)}>
                                                    <div className={css(styles.flipCardInner)}>
                                                      <div className={css(styles.flipCardFront)} style={frontstyle}>
                                                        <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                          className={css(styles.flipCardImageSizeMedium)} />
                                                        <div className={css(styles.flipCardTitleMediumMedium, styles.flipCardTitleMediumMediumRight)}>{item.Title}</div>
                                                      </div>
                                                      <div className={css(styles.flipCardBack)} style={backstyle}>
                                                        <p className={css(styles.flipCardBackTextMedium, styles.flipCardTitleMediumMediumRight)}>
                                                          {item.BackText}
                                                        </p>
                                                      </div>
                                                    </div>
                                                  </div>

                                                ) : CardSize == "Medium" && fonttextalign == "AlignLeft" && fontsizeop == "large" ?
                                                  (
                                                    <div className={css(styles.flipCardMedium)}>
                                                      <div className={css(styles.flipCardInner)}>
                                                        <div className={css(styles.flipCardFront)} style={frontstyle}>
                                                          <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                            className={css(styles.flipCardImageSizeMedium)} />
                                                          <div className={css(styles.flipCardTitleMediumMedium, styles.flipCardTitleMediumLargeLeft)}>{item.Title}</div>
                                                        </div>
                                                        <div className={css(styles.flipCardBack)} style={backstyle}>
                                                          <p className={css(styles.flipCardBackTextMedium, styles.flipCardTitleMediumLargeLeft)}>
                                                            {item.BackText}
                                                          </p>
                                                        </div>
                                                      </div>
                                                    </div>

                                                  ) : CardSize == "Medium" && fonttextalign == "AlignCenter" && fontsizeop == "large" ?
                                                    (
                                                      <div className={css(styles.flipCardMedium)}>
                                                        <div className={css(styles.flipCardInner)}>
                                                          <div className={css(styles.flipCardFront)} style={frontstyle}>
                                                            <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                              className={css(styles.flipCardImageSizeMedium)} />
                                                            <div className={css(styles.flipCardTitleMediumMedium, styles.flipCardTitleMediumLargeCenter)}>{item.Title}</div>
                                                          </div>
                                                          <div className={css(styles.flipCardBack)} style={backstyle}>
                                                            <p className={css(styles.flipCardBackTextMedium, styles.flipCardTitleMediumLargeCenter)}>
                                                              {item.BackText}
                                                            </p>
                                                          </div>
                                                        </div>
                                                      </div>

                                                    ) : CardSize == "Medium" && fonttextalign == "AlignRight" && fontsizeop == "large" ?
                                                      (
                                                        <div className={css(styles.flipCardMedium)}>
                                                          <div className={css(styles.flipCardInner)}>
                                                            <div className={css(styles.flipCardFront)} style={frontstyle}>
                                                              <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                                className={css(styles.flipCardImageSizeMedium)} />
                                                              <div className={css(styles.flipCardTitleMediumMedium, styles.flipCardTitleMediumLargeRight)}>{item.Title}</div>
                                                            </div>
                                                            <div className={css(styles.flipCardBack)} style={backstyle}>
                                                              <p className={css(styles.flipCardBackTextMedium, styles.flipCardTitleMediumLargeRight)}>
                                                                {item.BackText}
                                                              </p>
                                                            </div>
                                                          </div>
                                                        </div>

                                                      ) : CardSize == "Large" && fonttextalign == "AlignLeft" && fontsizeop == "small" ?
                                                        (
                                                          <div className={css(styles.flipCardLarge)}>
                                                            <div className={css(styles.flipCardInner)}>
                                                              <div className={css(styles.flipCardFront)} style={frontstyle}>
                                                                <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                                  className={css(styles.flipCardImageSizeLarge)} />
                                                                <div className={css(styles.flipCardTitleLargeSmall, styles.flipCardTitleLargeSmallLeft)}>{item.Title}</div>
                                                              </div>
                                                              <div className={css(styles.flipCardBack)} style={backstyle}>
                                                                <p className={css(styles.flipCardBackTextLarge, styles.flipCardTitleLargeSmallLeft)}>
                                                                  {item.BackText}
                                                                </p>
                                                              </div>
                                                            </div>
                                                          </div>

                                                        ) : CardSize == "Large" && fonttextalign == "AlignCenter" && fontsizeop == "small" ?
                                                          (
                                                            <div className={css(styles.flipCardLarge)}>
                                                              <div className={css(styles.flipCardInner)}>
                                                                <div className={css(styles.flipCardFront)} style={frontstyle}>
                                                                  <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                                    className={css(styles.flipCardImageSizeLarge)} />
                                                                  <div className={css(styles.flipCardTitleLargeSmall, styles.flipCardTitleLargeSmallCenter)}>{item.Title}</div>
                                                                </div>
                                                                <div className={css(styles.flipCardBack)} style={backstyle}>
                                                                  <p className={css(styles.flipCardBackTextLarge, styles.flipCardTitleLargeSmallCenter)}>
                                                                    {item.BackText}
                                                                  </p>
                                                                </div>
                                                              </div>
                                                            </div>

                                                          ) : CardSize == "Large" && fonttextalign == "AlignRight" && fontsizeop == "small" ?
                                                            (
                                                              <div className={css(styles.flipCardLarge)}>
                                                                <div className={css(styles.flipCardInner)}>
                                                                  <div className={css(styles.flipCardFront)} style={frontstyle}>
                                                                    <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                                      className={css(styles.flipCardImageSizeLarge)} />
                                                                    <div className={css(styles.flipCardTitleLargeSmall, styles.flipCardTitleLargeSmallRight)}>{item.Title}</div>
                                                                  </div>
                                                                  <div className={css(styles.flipCardBack)} style={backstyle}>
                                                                    <p className={css(styles.flipCardBackTextLarge, styles.flipCardTitleLargeSmallRight)}>
                                                                      {item.BackText}
                                                                    </p>
                                                                  </div>
                                                                </div>
                                                              </div>

                                                            ) : CardSize == "Large" && fonttextalign == "AlignLeft" && fontsizeop == "medium" ?
                                                              (
                                                                <div className={css(styles.flipCardLarge)}>
                                                                  <div className={css(styles.flipCardInner)}>
                                                                    <div className={css(styles.flipCardFront)} style={frontstyle}>
                                                                      <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                                        className={css(styles.flipCardImageSizeLarge)} />
                                                                      <div className={css(styles.flipCardTitleLargeMedium, styles.flipCardTitleLargeMediumLeft)}>{item.Title}</div>
                                                                    </div>
                                                                    <div className={css(styles.flipCardBack)} style={backstyle}>
                                                                      <p className={css(styles.flipCardBackTextLarge, styles.flipCardTitleLargeMediumLeft)}>
                                                                        {item.BackText}
                                                                      </p>
                                                                    </div>
                                                                  </div>
                                                                </div>

                                                              ) : CardSize == "Large" && fonttextalign == "AlignCenter" && fontsizeop == "medium" ?
                                                                (
                                                                  <div className={css(styles.flipCardLarge)}>
                                                                    <div className={css(styles.flipCardInner)}>
                                                                      <div className={css(styles.flipCardFront)} style={frontstyle}>
                                                                        <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                                          className={css(styles.flipCardImageSizeLarge)} />
                                                                        <div className={css(styles.flipCardTitleLargeMedium, styles.flipCardTitleLargeMediumCenter)}>{item.Title}</div>
                                                                      </div>
                                                                      <div className={css(styles.flipCardBack)} style={backstyle}>
                                                                        <p className={css(styles.flipCardBackTextLarge, styles.flipCardTitleLargeMediumCenter)}>
                                                                          {item.BackText}
                                                                        </p>
                                                                      </div>
                                                                    </div>
                                                                  </div>

                                                                ) : CardSize == "Large" && fonttextalign == "AlignRight" && fontsizeop == "medium" ?
                                                                  (
                                                                    <div className={css(styles.flipCardLarge)}>
                                                                      <div className={css(styles.flipCardInner)}>
                                                                        <div className={css(styles.flipCardFront)} style={frontstyle}>
                                                                          <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                                            className={css(styles.flipCardImageSizeLarge)} />
                                                                          <div className={css(styles.flipCardTitleLargeMedium, styles.flipCardTitleLargeMediumRight)}>{item.Title}</div>
                                                                        </div>
                                                                        <div className={css(styles.flipCardBack)} style={backstyle}>
                                                                          <p className={css(styles.flipCardBackTextLarge, styles.flipCardTitleLargeMediumRight)}>
                                                                            {item.BackText}
                                                                          </p>
                                                                        </div>
                                                                      </div>
                                                                    </div>

                                                                  ) : CardSize == "Large" && fonttextalign == "AlignLeft" && fontsizeop == "large" ?
                                                                    (
                                                                      <div className={css(styles.flipCardLarge)}>
                                                                        <div className={css(styles.flipCardInner)}>
                                                                          <div className={css(styles.flipCardFront)} style={frontstyle}>
                                                                            <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                                              className={css(styles.flipCardImageSizeLarge)} />
                                                                            <div className={css(styles.flipCardTitleLargeLarge, styles.flipCardTitleLargeLargeLeft)}>{item.Title}</div>
                                                                          </div>
                                                                          <div className={css(styles.flipCardBack)} style={backstyle}>
                                                                            <p className={css(styles.flipCardBackTextLarge, styles.flipCardTitleLargeLargeLeft)}>
                                                                              {item.BackText}
                                                                            </p>
                                                                          </div>
                                                                        </div>
                                                                      </div>

                                                                    ) : CardSize == "Large" && fonttextalign == "AlignCenter" && fontsizeop == "large" ?
                                                                      (
                                                                        <div className={css(styles.flipCardLarge)}>
                                                                          <div className={css(styles.flipCardInner)}>
                                                                            <div className={css(styles.flipCardFront)} style={frontstyle}>
                                                                              <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                                                className={css(styles.flipCardImageSizeLarge)} />
                                                                              <div className={css(styles.flipCardTitleLargeLarge, styles.flipCardTitleLargeLargeCenter)}>{item.Title}</div>
                                                                            </div>
                                                                            <div className={css(styles.flipCardBack)} style={backstyle}>
                                                                              <p className={css(styles.flipCardBackTextLarge, styles.flipCardTitleLargeLargeCenter)}>
                                                                                {item.BackText}
                                                                              </p>
                                                                            </div>
                                                                          </div>
                                                                        </div>

                                                                      ) : CardSize == "Large" && fonttextalign == "AlignRight" && fontsizeop == "large" ?
                                                                        (
                                                                          <div className={css(styles.flipCardLarge)}>
                                                                            <div className={css(styles.flipCardInner)}>
                                                                              <div className={css(styles.flipCardFront)} style={frontstyle}>
                                                                                <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                                                  className={css(styles.flipCardImageSizeLarge)} />
                                                                                <div className={css(styles.flipCardTitleLargeLarge, styles.flipCardTitleLargeLargeRight)}>{item.Title}</div>
                                                                              </div>
                                                                              <div className={css(styles.flipCardBack)} style={backstyle}>
                                                                                <p className={css(styles.flipCardBackTextLarge, styles.flipCardTitleLargeLargeRight)}>
                                                                                  {item.BackText}
                                                                                </p>
                                                                              </div>
                                                                            </div>
                                                                          </div>

                                                                        ) : CardSize == "ButtonSmall" && fonttextalign == "AlignLeft" && fontsizeop == "small" ?
                                                                          (
                                                                            <div className={css(styles.flipCardButtonSmall)}>
                                                                              <div className={css(styles.flipCardInner)}>
                                                                                <div className={css(styles.flipCardFront)} style={frontstyle}>

                                                                                  <div className={css(styles.flipCardTitleButtonSmall, styles.flipCardTitleButtonSmallLeft)}>{item.Title}</div>
                                                                                </div>
                                                                                <div className={css(styles.flipCardBack)} style={backstyle}>
                                                                                  <p className={css(styles.flipCardBackTextButton, styles.flipCardTitleButtonSmallLeft)}>
                                                                                    {item.BackText}
                                                                                  </p>
                                                                                </div>
                                                                              </div>
                                                                            </div>

                                                                          ) : CardSize == "ButtonSmall" && fonttextalign == "AlignCenter" && fontsizeop == "small" ?
                                                                            (
                                                                              <div className={css(styles.flipCardButtonSmall)}>
                                                                                <div className={css(styles.flipCardInner)}>
                                                                                  <div className={css(styles.flipCardFront)} style={frontstyle}>

                                                                                    <div className={css(styles.flipCardTitleButtonSmall, styles.flipCardTitleButtonSmallCenter)}>{item.Title}</div>
                                                                                  </div>
                                                                                  <div className={css(styles.flipCardBack)} style={backstyle}>
                                                                                    <p className={css(styles.flipCardBackTextButton, styles.flipCardTitleButtonSmallCenter)}>
                                                                                      {item.BackText}
                                                                                    </p>
                                                                                  </div>
                                                                                </div>
                                                                              </div>

                                                                            ) : CardSize == "ButtonSmall" && fonttextalign == "AlignRight" && fontsizeop == "small" ?
                                                                              (
                                                                                <div className={css(styles.flipCardButtonSmall)}>
                                                                                  <div className={css(styles.flipCardInner)}>
                                                                                    <div className={css(styles.flipCardFront)} style={frontstyle}>

                                                                                      <div className={css(styles.flipCardTitleButtonSmall, styles.flipCardTitleButtonSmallRight)}>{item.Title}</div>
                                                                                    </div>
                                                                                    <div className={css(styles.flipCardBack)} style={backstyle}>
                                                                                      <p className={css(styles.flipCardBackTextButton, styles.flipCardTitleButtonSmallRight)}>
                                                                                        {item.BackText}
                                                                                      </p>
                                                                                    </div>
                                                                                  </div>
                                                                                </div>

                                                                              ) : CardSize == "ButtonSmall" && fonttextalign == "AlignLeft" && fontsizeop == "medium" ?
                                                                                (
                                                                                  <div className={css(styles.flipCardButtonSmall)}>
                                                                                    <div className={css(styles.flipCardInner)}>
                                                                                      <div className={css(styles.flipCardFront)} style={frontstyle}>

                                                                                        <div className={css(styles.flipCardTitleButtonSmall, styles.flipCardTitleButtonSmallMediumLeft)}>{item.Title}</div>
                                                                                      </div>
                                                                                      <div className={css(styles.flipCardBack)} style={backstyle}>
                                                                                        <p className={css(styles.flipCardBackTextButtonMedium, styles.flipCardTitleCompactMediumLeft)}>
                                                                                          {item.BackText}
                                                                                        </p>
                                                                                      </div>
                                                                                    </div>
                                                                                  </div>

                                                                                ) : CardSize == "ButtonSmall" && fonttextalign == "AlignCenter" && fontsizeop == "medium" ?
                                                                                  (
                                                                                    <div className={css(styles.flipCardButtonSmall)}>
                                                                                      <div className={css(styles.flipCardInner)}>
                                                                                        <div className={css(styles.flipCardFront)} style={frontstyle}>

                                                                                          <div className={css(styles.flipCardTitleButtonSmall, styles.flipCardTitleButtonSmallMediumCenter)}>{item.Title}</div>
                                                                                        </div>
                                                                                        <div className={css(styles.flipCardBack)} style={backstyle}>
                                                                                          <p className={css(styles.flipCardBackTextButtonMedium, styles.flipCardTitleCompactMediumCenter)}>
                                                                                            {item.BackText}
                                                                                          </p>
                                                                                        </div>
                                                                                      </div>
                                                                                    </div>

                                                                                  ) : CardSize == "ButtonSmall" && fonttextalign == "AlignRight" && fontsizeop == "medium" ?
                                                                                    (
                                                                                      <div className={css(styles.flipCardButtonSmall)}>
                                                                                        <div className={css(styles.flipCardInner)}>
                                                                                          <div className={css(styles.flipCardFront)} style={frontstyle}>

                                                                                            <div className={css(styles.flipCardTitleButtonSmall, styles.flipCardTitleButtonSmallMediumRight)}>{item.Title}</div>
                                                                                          </div>
                                                                                          <div className={css(styles.flipCardBack)} style={backstyle}>
                                                                                            <p className={css(styles.flipCardBackTextButtonMedium, styles.flipCardTitleCompactMediumRight)}>
                                                                                              {item.BackText}
                                                                                            </p>
                                                                                          </div>
                                                                                        </div>
                                                                                      </div>

                                                                                    ) : CardSize == "ButtonSmall" && fonttextalign == "AlignLeft" && fontsizeop == "large" ?
                                                                                      (
                                                                                        <div className={css(styles.flipCardButtonSmall)}>
                                                                                          <div className={css(styles.flipCardInner)}>
                                                                                            <div className={css(styles.flipCardFront)} style={frontstyle}>

                                                                                              <div className={css(styles.flipCardTitleButtonSmall, styles.flipCardTitleButtonSmallLargeLeft)}>{item.Title}</div>
                                                                                            </div>
                                                                                            <div className={css(styles.flipCardBack)} style={backstyle}>
                                                                                              <p className={css(styles.flipCardBackTextButtonLarge, styles.flipCardTitleCompactLargeLeft)}>
                                                                                                {item.BackText}
                                                                                              </p>
                                                                                            </div>
                                                                                          </div>
                                                                                        </div>

                                                                                      ) : CardSize == "ButtonSmall" && fonttextalign == "AlignCenter" && fontsizeop == "large" ?
                                                                                        (
                                                                                          <div className={css(styles.flipCardButtonSmall)}>
                                                                                            <div className={css(styles.flipCardInner)}>
                                                                                              <div className={css(styles.flipCardFront)} style={frontstyle}>

                                                                                                <div className={css(styles.flipCardTitleButtonSmall, styles.flipCardTitleButtonSmallLargeCenter)}>{item.Title}</div>
                                                                                              </div>
                                                                                              <div className={css(styles.flipCardBack)} style={backstyle}>
                                                                                                <p className={css(styles.flipCardBackTextButtonLarge, styles.flipCardTitleButtonSmallLargeCenter)}>
                                                                                                  {item.BackText}
                                                                                                </p>
                                                                                              </div>
                                                                                            </div>
                                                                                          </div>

                                                                                        ) : CardSize == "ButtonSmall" && fonttextalign == "AlignRight" && fontsizeop == "large" ?
                                                                                          (
                                                                                            <div className={css(styles.flipCardButtonSmall)}>
                                                                                              <div className={css(styles.flipCardInner)}>
                                                                                                <div className={css(styles.flipCardFront)} style={frontstyle}>

                                                                                                  <div className={css(styles.flipCardTitleButtonSmall, styles.flipCardTitleButtonSmallLargeRight)}>{item.Title}</div>
                                                                                                </div>
                                                                                                <div className={css(styles.flipCardBack)} style={backstyle}>
                                                                                                  <p className={css(styles.flipCardBackTextButtonLarge, styles.flipCardTitleButtonSmallLargeRight)}>
                                                                                                    {item.BackText}
                                                                                                  </p>
                                                                                                </div>
                                                                                              </div>
                                                                                            </div>

                                                                                          ) : CardSize == "ButtonSmall" && fonttextalign == "AlignLeft" && fontsizeop == "small" ?
                                                                                            (
                                                                                              <div className={css(styles.flipCardButtonSmall)}>
                                                                                                <div className={css(styles.flipCardInner)}>
                                                                                                  <div className={css(styles.flipCardFront)} style={frontstyle}>

                                                                                                    <div className={css(styles.flipCardTitleButtonSmall, styles.flipCardTitleButtonSmallLeft)}>{item.Title}</div>
                                                                                                  </div>
                                                                                                  <div className={css(styles.flipCardBack)} style={backstyle}>
                                                                                                    <p className={css(styles.flipCardBackTextButton, styles.flipCardTitleButtonSmallLeft)}>
                                                                                                      {item.BackText}
                                                                                                    </p>
                                                                                                  </div>
                                                                                                </div>
                                                                                              </div>

                                                                                            ) : CardSize == "ButtonSmall" && fonttextalign == "AlignCenter" && fontsizeop == "small" ?
                                                                                              (
                                                                                                <div className={css(styles.flipCardButtonSmall)}>
                                                                                                  <div className={css(styles.flipCardInner)}>
                                                                                                    <div className={css(styles.flipCardFront)} style={frontstyle}>

                                                                                                      <div className={css(styles.flipCardTitleButtonSmall, styles.flipCardTitleButtonSmallCenter)}>{item.Title}</div>
                                                                                                    </div>
                                                                                                    <div className={css(styles.flipCardBack)} style={backstyle}>
                                                                                                      <p className={css(styles.flipCardBackTextButton, styles.flipCardTitleButtonSmallCenter)}>
                                                                                                        {item.BackText}
                                                                                                      </p>
                                                                                                    </div>
                                                                                                  </div>
                                                                                                </div>

                                                                                              ) : false

                )

                : enableFlipFlop == true && flipCardAxis == "X" ?
                  (
                    CardSize == "Compact" && fonttextalign == "AlignLeft" && fontsizeop == "small" ?
                      (
                        <div className={css(styles.flipCardX)}>
                          <div className={css(styles.flipCardInner)}>
                            <div className={css(styles.flipCardFront)} style={frontstyle}>
                              <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                className={css(styles.flipCardImageSizeCompact)} />
                              <div className={css(styles.flipCardTitleCompactSmall, styles.flipCardTitleCompactSmallLeft)}>{item.Title}</div>
                            </div>
                            <div className={css(styles.flipCardXBack)} style={backstyle}>
                              <p className={css(styles.flipCardBackTextCompact, styles.flipCardTitleCompactSmallLeft)}>
                                {item.BackText}
                              </p>
                            </div>
                          </div>
                        </div>

                      ) : CardSize == "Compact" && fonttextalign == "AlignCenter" && fontsizeop == "small" ?
                        (
                          <div className={css(styles.flipCardX)}>
                            <div className={css(styles.flipCardInner)}>
                              <div className={css(styles.flipCardFront)} style={frontstyle}>
                                <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                  className={css(styles.flipCardImageSizeCompact)} />
                                <div className={css(styles.flipCardTitleCompactSmall, styles.flipCardTitleCompactSmallCenter)}>{item.Title}</div>
                              </div>
                              <div className={css(styles.flipCardXBack)} style={backstyle}>
                                <p className={css(styles.flipCardBackTextCompact, styles.flipCardTitleCompactSmallCenter)}>
                                  {item.BackText}
                                </p>
                              </div>
                            </div>
                          </div>

                        ) : CardSize == "Compact" && fonttextalign == "AlignRight" && fontsizeop == "small" ?
                          (
                            <div className={css(styles.flipCardX)}>
                              <div className={css(styles.flipCardInner)}>
                                <div className={css(styles.flipCardFront)} style={frontstyle}>
                                  <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                    className={css(styles.flipCardImageSizeCompact)} />
                                  <div className={css(styles.flipCardTitleCompactSmall, styles.flipCardTitleCompactSmallRight)}>{item.Title}</div>
                                </div>
                                <div className={css(styles.flipCardXBack)} style={backstyle}>
                                  <p className={css(styles.flipCardBackTextCompact, styles.flipCardTitleCompactSmallRight)}>
                                    {item.BackText}
                                  </p>
                                </div>
                              </div>
                            </div>

                          ) : CardSize == "Compact" && fonttextalign == "AlignLeft" && fontsizeop == "medium" ?
                            (
                              <div className={css(styles.flipCardX)}>
                                <div className={css(styles.flipCardInner)}>
                                  <div className={css(styles.flipCardFront)} style={frontstyle}>
                                    <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                      className={css(styles.flipCardImageSizeCompact)} />
                                    <div className={css(styles.flipCardTitleCompactMedium, styles.flipCardTitleCompactMediumLeft)}>{item.Title}</div>
                                  </div>
                                  <div className={css(styles.flipCardXBack)} style={backstyle}>
                                    <p className={css(styles.flipCardBackTextCompact, styles.flipCardTitleCompactMediumLeft)}>
                                      {item.BackText}
                                    </p>
                                  </div>
                                </div>
                              </div>

                            ) : CardSize == "Compact" && fonttextalign == "AlignCenter" && fontsizeop == "medium" ?
                              (
                                <div className={css(styles.flipCardX)}>
                                  <div className={css(styles.flipCardInner)}>
                                    <div className={css(styles.flipCardFront)} style={frontstyle}>
                                      <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                        className={css(styles.flipCardImageSizeCompact)} />
                                      <div className={css(styles.flipCardTitleCompactMedium, styles.flipCardTitleCompactMediumCenter)}>{item.Title}</div>
                                    </div>
                                    <div className={css(styles.flipCardXBack)} style={backstyle}>
                                      <p className={css(styles.flipCardBackTextCompact, styles.flipCardTitleCompactMediumCenter)}>
                                        {item.BackText}
                                      </p>
                                    </div>
                                  </div>
                                </div>

                              ) : CardSize == "Compact" && fonttextalign == "AlignRight" && fontsizeop == "medium" ?
                                (
                                  <div className={css(styles.flipCardX)}>
                                    <div className={css(styles.flipCardInner)}>
                                      <div className={css(styles.flipCardFront)} style={frontstyle}>
                                        <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                          className={css(styles.flipCardImageSizeCompact)} />
                                        <div className={css(styles.flipCardTitleCompactMedium, styles.flipCardTitleCompactMediumRight)}>{item.Title}</div>
                                      </div>
                                      <div className={css(styles.flipCardXBack)} style={backstyle}>
                                        <p className={css(styles.flipCardBackTextCompact, styles.flipCardTitleCompactMediumRight)}>
                                          {item.BackText}
                                        </p>
                                      </div>
                                    </div>
                                  </div>

                                ) : CardSize == "Compact" && fonttextalign == "AlignLeft" && fontsizeop == "large" ?
                                  (
                                    <div className={css(styles.flipCardX)}>
                                      <div className={css(styles.flipCardInner)}>
                                        <div className={css(styles.flipCardFront)} style={frontstyle}>
                                          <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                            className={css(styles.flipCardImageSizeCompact)} />
                                          <div className={css(styles.flipCardTitleCompactLarge, styles.flipCardTitleCompactLargeLeft)}>{item.Title}</div>
                                        </div>
                                        <div className={css(styles.flipCardXBack)} style={backstyle}>
                                          <p className={css(styles.flipCardBackTextCompact, styles.flipCardTitleCompactLargeLeft)}>
                                            {item.BackText}
                                          </p>
                                        </div>
                                      </div>
                                    </div>

                                  ) : CardSize == "Compact" && fonttextalign == "AlignCenter" && fontsizeop == "large" ?
                                    (
                                      <div className={css(styles.flipCardX)}>
                                        <div className={css(styles.flipCardInner)}>
                                          <div className={css(styles.flipCardFront)} style={frontstyle}>
                                            <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                              className={css(styles.flipCardImageSizeCompact)} />
                                            <div className={css(styles.flipCardTitleCompactLarge, styles.flipCardTitleCompactLargeCenter)}>{item.Title}</div>
                                          </div>
                                          <div className={css(styles.flipCardXBack)} style={backstyle}>
                                            <p className={css(styles.flipCardBackTextCompact, styles.flipCardTitleCompactLargeCenter)}>
                                              {item.BackText}
                                            </p>
                                          </div>
                                        </div>
                                      </div>

                                    ) : CardSize == "Compact" && fonttextalign == "AlignRight" && fontsizeop == "large" ?
                                      (
                                        <div className={css(styles.flipCardX)}>
                                          <div className={css(styles.flipCardInner)}>
                                            <div className={css(styles.flipCardFront)} style={frontstyle}>
                                              <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                className={css(styles.flipCardImageSizeCompact)} />
                                              <div className={css(styles.flipCardTitleCompactLarge, styles.flipCardTitleCompactLargeRight)}>{item.Title}</div>
                                            </div>
                                            <div className={css(styles.flipCardXBack)} style={backstyle}>
                                              <p className={css(styles.flipCardBackTextCompact, styles.flipCardTitleCompactLargeRight)}>
                                                {item.BackText}
                                              </p>
                                            </div>
                                          </div>
                                        </div>

                                      ) : CardSize == "Medium" && fonttextalign == "AlignLeft" && fontsizeop == "small" ?
                                        (
                                          <div className={css(styles.flipCardXMedium)}>
                                            <div className={css(styles.flipCardInner)}>
                                              <div className={css(styles.flipCardFront)} style={frontstyle}>
                                                <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                  className={css(styles.flipCardImageSizeMedium)} />
                                                <div className={css(styles.flipCardTitleMediumSmall, styles.flipCardTitleMediumSmallLeft)}>{item.Title}</div>
                                              </div>
                                              <div className={css(styles.flipCardXBack)} style={backstyle}>
                                                <p className={css(styles.flipCardBackTextMedium, styles.flipCardTitleMediumSmallLeft)}>
                                                  {item.BackText}
                                                </p>
                                              </div>
                                            </div>
                                          </div>

                                        ) : CardSize == "Medium" && fonttextalign == "AlignCenter" && fontsizeop == "small" ?
                                          (
                                            <div className={css(styles.flipCardXMedium)}>
                                              <div className={css(styles.flipCardInner)}>
                                                <div className={css(styles.flipCardFront)} style={frontstyle}>
                                                  <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                    className={css(styles.flipCardImageSizeMedium)} />
                                                  <div className={css(styles.flipCardTitleMediumSmall, styles.flipCardTitleMediumSmallCenter)}>{item.Title}</div>
                                                </div>
                                                <div className={css(styles.flipCardXBack)} style={backstyle}>
                                                  <p className={css(styles.flipCardBackTextMedium, styles.flipCardTitleMediumSmallCenter)}>
                                                    {item.BackText}
                                                  </p>
                                                </div>
                                              </div>
                                            </div>

                                          ) : CardSize == "Medium" && fonttextalign == "AlignRight" && fontsizeop == "small" ?
                                            (
                                              <div className={css(styles.flipCardXMedium)}>
                                                <div className={css(styles.flipCardInner)}>
                                                  <div className={css(styles.flipCardFront)} style={frontstyle}>
                                                    <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                      className={css(styles.flipCardImageSizeMedium)} />
                                                    <div className={css(styles.flipCardTitleMediumSmall, styles.flipCardTitleMediumSmallRight)}>{item.Title}</div>
                                                  </div>
                                                  <div className={css(styles.flipCardXBack)} style={backstyle}>
                                                    <p className={css(styles.flipCardBackTextMedium, styles.flipCardTitleMediumSmallRight)}>
                                                      {item.BackText}
                                                    </p>
                                                  </div>
                                                </div>
                                              </div>

                                            ) : CardSize == "Medium" && fonttextalign == "AlignLeft" && fontsizeop == "medium" ?
                                              (
                                                <div className={css(styles.flipCardXMedium)}>
                                                  <div className={css(styles.flipCardInner)}>
                                                    <div className={css(styles.flipCardFront)} style={frontstyle}>
                                                      <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                        className={css(styles.flipCardImageSizeMedium)} />
                                                      <div className={css(styles.flipCardTitleMediumMedium, styles.flipCardTitleMediumMediumLeft)}>{item.Title}</div>
                                                    </div>
                                                    <div className={css(styles.flipCardXBack)} style={backstyle}>
                                                      <p className={css(styles.flipCardBackTextMedium, styles.flipCardTitleMediumMediumLeft)}>
                                                        {item.BackText}
                                                      </p>
                                                    </div>
                                                  </div>
                                                </div>

                                              ) : CardSize == "Medium" && fonttextalign == "AlignCenter" && fontsizeop == "medium" ?
                                                (
                                                  <div className={css(styles.flipCardXMedium)}>
                                                    <div className={css(styles.flipCardInner)}>
                                                      <div className={css(styles.flipCardFront)} style={frontstyle}>
                                                        <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                          className={css(styles.flipCardImageSizeMedium)} />
                                                        <div className={css(styles.flipCardTitleMediumMedium, styles.flipCardTitleMediumMediumCenter)}>{item.Title}</div>
                                                      </div>
                                                      <div className={css(styles.flipCardXBack)} style={backstyle}>
                                                        <p className={css(styles.flipCardBackTextMedium, styles.flipCardTitleMediumMediumCenter)}>
                                                          {item.BackText}
                                                        </p>
                                                      </div>
                                                    </div>
                                                  </div>

                                                ) : CardSize == "Medium" && fonttextalign == "AlignRight" && fontsizeop == "medium" ?
                                                  (
                                                    <div className={css(styles.flipCardXMedium)}>
                                                      <div className={css(styles.flipCardInner)}>
                                                        <div className={css(styles.flipCardFront)} style={frontstyle}>
                                                          <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                            className={css(styles.flipCardImageSizeMedium)} />
                                                          <div className={css(styles.flipCardTitleMediumMedium, styles.flipCardTitleMediumMediumRight)}>{item.Title}</div>
                                                        </div>
                                                        <div className={css(styles.flipCardXBack)} style={backstyle}>
                                                          <p className={css(styles.flipCardBackTextMedium, styles.flipCardTitleMediumMediumRight)}>
                                                            {item.BackText}
                                                          </p>
                                                        </div>
                                                      </div>
                                                    </div>

                                                  ) : CardSize == "Medium" && fonttextalign == "AlignLeft" && fontsizeop == "large" ?
                                                    (
                                                      <div className={css(styles.flipCardXMedium)}>
                                                        <div className={css(styles.flipCardInner)}>
                                                          <div className={css(styles.flipCardFront)} style={frontstyle}>
                                                            <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                              className={css(styles.flipCardImageSizeMedium)} />
                                                            <div className={css(styles.flipCardTitleMediumMedium, styles.flipCardTitleMediumLargeLeft)}>{item.Title}</div>
                                                          </div>
                                                          <div className={css(styles.flipCardXBack)} style={backstyle}>
                                                            <p className={css(styles.flipCardBackTextMedium, styles.flipCardTitleMediumLargeLeft)}>
                                                              {item.BackText}
                                                            </p>
                                                          </div>
                                                        </div>
                                                      </div>

                                                    ) : CardSize == "Medium" && fonttextalign == "AlignCenter" && fontsizeop == "large" ?
                                                      (
                                                        <div className={css(styles.flipCardXMedium)}>
                                                          <div className={css(styles.flipCardInner)}>
                                                            <div className={css(styles.flipCardFront)} style={frontstyle}>
                                                              <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                                className={css(styles.flipCardImageSizeMedium)} />
                                                              <div className={css(styles.flipCardTitleMediumMedium, styles.flipCardTitleMediumLargeCenter)}>{item.Title}</div>
                                                            </div>
                                                            <div className={css(styles.flipCardXBack)} style={backstyle}>
                                                              <p className={css(styles.flipCardBackTextMedium, styles.flipCardTitleMediumLargeCenter)}>
                                                                {item.BackText}
                                                              </p>
                                                            </div>
                                                          </div>
                                                        </div>

                                                      ) : CardSize == "Medium" && fonttextalign == "AlignRight" && fontsizeop == "large" ?
                                                        (
                                                          <div className={css(styles.flipCardXMedium)}>
                                                            <div className={css(styles.flipCardInner)}>
                                                              <div className={css(styles.flipCardFront)} style={frontstyle}>
                                                                <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                                  className={css(styles.flipCardImageSizeMedium)} />
                                                                <div className={css(styles.flipCardTitleMediumMedium, styles.flipCardTitleMediumLargeRight)}>{item.Title}</div>
                                                              </div>
                                                              <div className={css(styles.flipCardXBack)} style={backstyle}>
                                                                <p className={css(styles.flipCardBackTextMedium, styles.flipCardTitleMediumLargeRight)}>
                                                                  {item.BackText}
                                                                </p>
                                                              </div>
                                                            </div>
                                                          </div>

                                                        ) : CardSize == "Large" && fonttextalign == "AlignLeft" && fontsizeop == "small" ?
                                                          (
                                                            <div className={css(styles.flipCardXLarge)}>
                                                              <div className={css(styles.flipCardInner)}>
                                                                <div className={css(styles.flipCardFront)} style={frontstyle}>
                                                                  <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                                    className={css(styles.flipCardImageSizeLarge)} />
                                                                  <div className={css(styles.flipCardTitleLargeSmall, styles.flipCardTitleLargeSmallLeft)}>{item.Title}</div>
                                                                </div>
                                                                <div className={css(styles.flipCardXBack)} style={backstyle}>
                                                                  <p className={css(styles.flipCardBackTextLarge, styles.flipCardTitleLargeSmallLeft)}>
                                                                    {item.BackText}
                                                                  </p>
                                                                </div>
                                                              </div>
                                                            </div>

                                                          ) : CardSize == "Large" && fonttextalign == "AlignCenter" && fontsizeop == "small" ?
                                                            (
                                                              <div className={css(styles.flipCardXLarge)}>
                                                                <div className={css(styles.flipCardInner)}>
                                                                  <div className={css(styles.flipCardFront)} style={frontstyle}>
                                                                    <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                                      className={css(styles.flipCardImageSizeLarge)} />
                                                                    <div className={css(styles.flipCardTitleLargeSmall, styles.flipCardTitleLargeSmallCenter)}>{item.Title}</div>
                                                                  </div>
                                                                  <div className={css(styles.flipCardXBack)} style={backstyle}>
                                                                    <p className={css(styles.flipCardBackTextLarge, styles.flipCardTitleLargeSmallCenter)}>
                                                                      {item.BackText}
                                                                    </p>
                                                                  </div>
                                                                </div>
                                                              </div>

                                                            ) : CardSize == "Large" && fonttextalign == "AlignRight" && fontsizeop == "small" ?
                                                              (
                                                                <div className={css(styles.flipCardXLarge)}>
                                                                  <div className={css(styles.flipCardInner)}>
                                                                    <div className={css(styles.flipCardFront)} style={frontstyle}>
                                                                      <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                                        className={css(styles.flipCardImageSizeLarge)} />
                                                                      <div className={css(styles.flipCardTitleLargeSmall, styles.flipCardTitleLargeSmallRight)}>{item.Title}</div>
                                                                    </div>
                                                                    <div className={css(styles.flipCardXBack)} style={backstyle}>
                                                                      <p className={css(styles.flipCardBackTextLarge, styles.flipCardTitleLargeSmallRight)}>
                                                                        {item.BackText}
                                                                      </p>
                                                                    </div>
                                                                  </div>
                                                                </div>

                                                              ) : CardSize == "Large" && fonttextalign == "AlignLeft" && fontsizeop == "medium" ?
                                                                (
                                                                  <div className={css(styles.flipCardXLarge)}>
                                                                    <div className={css(styles.flipCardInner)}>
                                                                      <div className={css(styles.flipCardFront)} style={frontstyle}>
                                                                        <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                                          className={css(styles.flipCardImageSizeLarge)} />
                                                                        <div className={css(styles.flipCardTitleLargeMedium, styles.flipCardTitleLargeMediumLeft)}>{item.Title}</div>
                                                                      </div>
                                                                      <div className={css(styles.flipCardXBack)} style={backstyle}>
                                                                        <p className={css(styles.flipCardBackTextLarge, styles.flipCardTitleLargeMediumLeft)}>
                                                                          {item.BackText}
                                                                        </p>
                                                                      </div>
                                                                    </div>
                                                                  </div>

                                                                ) : CardSize == "Large" && fonttextalign == "AlignCenter" && fontsizeop == "medium" ?
                                                                  (
                                                                    <div className={css(styles.flipCardXLarge)}>
                                                                      <div className={css(styles.flipCardInner)}>
                                                                        <div className={css(styles.flipCardFront)} style={frontstyle}>
                                                                          <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                                            className={css(styles.flipCardImageSizeLarge)} />
                                                                          <div className={css(styles.flipCardTitleLargeMedium, styles.flipCardTitleLargeMediumCenter)}>{item.Title}</div>
                                                                        </div>
                                                                        <div className={css(styles.flipCardXBack)} style={backstyle}>
                                                                          <p className={css(styles.flipCardBackTextLarge, styles.flipCardTitleLargeMediumCenter)}>
                                                                            {item.BackText}
                                                                          </p>
                                                                        </div>
                                                                      </div>
                                                                    </div>

                                                                  ) : CardSize == "Large" && fonttextalign == "AlignRight" && fontsizeop == "medium" ?
                                                                    (
                                                                      <div className={css(styles.flipCardXLarge)}>
                                                                        <div className={css(styles.flipCardInner)}>
                                                                          <div className={css(styles.flipCardFront)} style={frontstyle}>
                                                                            <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                                              className={css(styles.flipCardImageSizeLarge)} />
                                                                            <div className={css(styles.flipCardTitleLargeMedium, styles.flipCardTitleLargeMediumRight)}>{item.Title}</div>
                                                                          </div>
                                                                          <div className={css(styles.flipCardXBack)} style={backstyle}>
                                                                            <p className={css(styles.flipCardBackTextLarge, styles.flipCardTitleLargeMediumRight)}>
                                                                              {item.BackText}
                                                                            </p>
                                                                          </div>
                                                                        </div>
                                                                      </div>

                                                                    ) : CardSize == "Large" && fonttextalign == "AlignLeft" && fontsizeop == "large" ?
                                                                      (
                                                                        <div className={css(styles.flipCardXLarge)}>
                                                                          <div className={css(styles.flipCardInner)}>
                                                                            <div className={css(styles.flipCardFront)} style={frontstyle}>
                                                                              <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                                                className={css(styles.flipCardImageSizeLarge)} />
                                                                              <div className={css(styles.flipCardTitleLargeLarge, styles.flipCardTitleLargeLargeLeft)}>{item.Title}</div>
                                                                            </div>
                                                                            <div className={css(styles.flipCardXBack)} style={backstyle}>
                                                                              <p className={css(styles.flipCardBackTextLarge, styles.flipCardTitleLargeLargeLeft)}>
                                                                                {item.BackText}
                                                                              </p>
                                                                            </div>
                                                                          </div>
                                                                        </div>

                                                                      ) : CardSize == "Large" && fonttextalign == "AlignCenter" && fontsizeop == "large" ?
                                                                        (
                                                                          <div className={css(styles.flipCardXLarge)}>
                                                                            <div className={css(styles.flipCardInner)}>
                                                                              <div className={css(styles.flipCardFront)} style={frontstyle}>
                                                                                <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                                                  className={css(styles.flipCardImageSizeLarge)} />
                                                                                <div className={css(styles.flipCardTitleLargeLarge, styles.flipCardTitleLargeLargeCenter)}>{item.Title}</div>
                                                                              </div>
                                                                              <div className={css(styles.flipCardXBack)} style={backstyle}>
                                                                                <p className={css(styles.flipCardBackTextLarge, styles.flipCardTitleLargeLargeCenter)}>
                                                                                  {item.BackText}
                                                                                </p>
                                                                              </div>
                                                                            </div>
                                                                          </div>

                                                                        ) : CardSize == "Large" && fonttextalign == "AlignRight" && fontsizeop == "large" ?
                                                                          (
                                                                            <div className={css(styles.flipCardXLarge)}>
                                                                              <div className={css(styles.flipCardInner)}>
                                                                                <div className={css(styles.flipCardFront)} style={frontstyle}>
                                                                                  <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                                                    className={css(styles.flipCardImageSizeLarge)} />
                                                                                  <div className={css(styles.flipCardTitleLargeLarge, styles.flipCardTitleLargeLargeRight)}>{item.Title}</div>
                                                                                </div>
                                                                                <div className={css(styles.flipCardXBack)} style={backstyle}>
                                                                                  <p className={css(styles.flipCardBackTextLarge, styles.flipCardTitleLargeLargeRight)}>
                                                                                    {item.BackText}
                                                                                  </p>
                                                                                </div>
                                                                              </div>
                                                                            </div>

                                                                          ) : CardSize == "ButtonSmall" && fonttextalign == "AlignLeft" && fontsizeop == "small" ?
                                                                            (
                                                                              <div className={css(styles.flipCardButtonSmall)}>
                                                                                <div className={css(styles.flipCardInner)}>
                                                                                  <div className={css(styles.flipCardFront)} style={frontstyle}>

                                                                                    <div className={css(styles.flipCardTitleButtonSmall, styles.flipCardTitleButtonSmallLeft)}>{item.Title}</div>
                                                                                  </div>
                                                                                  <div className={css(styles.flipCardBack)} style={backstyle}>
                                                                                    <p className={css(styles.flipCardBackTextButton, styles.flipCardTitleButtonSmallLeft)}>
                                                                                      {item.BackText}
                                                                                    </p>
                                                                                  </div>
                                                                                </div>
                                                                              </div>

                                                                            ) : CardSize == "ButtonSmall" && fonttextalign == "AlignCenter" && fontsizeop == "small" ?
                                                                              (
                                                                                <div className={css(styles.flipCardButtonSmall)}>
                                                                                  <div className={css(styles.flipCardInner)}>
                                                                                    <div className={css(styles.flipCardFront)} style={frontstyle}>

                                                                                      <div className={css(styles.flipCardTitleButtonSmall, styles.flipCardTitleButtonSmallCenter)}>{item.Title}</div>
                                                                                    </div>
                                                                                    <div className={css(styles.flipCardBack)} style={backstyle}>
                                                                                      <p className={css(styles.flipCardBackTextButton, styles.flipCardTitleButtonSmallCenter)}>
                                                                                        {item.BackText}
                                                                                      </p>
                                                                                    </div>
                                                                                  </div>
                                                                                </div>

                                                                              ) : CardSize == "ButtonSmall" && fonttextalign == "AlignRight" && fontsizeop == "small" ?
                                                                                (
                                                                                  <div className={css(styles.flipCardButtonSmall)}>
                                                                                    <div className={css(styles.flipCardInner)}>
                                                                                      <div className={css(styles.flipCardFront)} style={frontstyle}>

                                                                                        <div className={css(styles.flipCardTitleButtonSmall, styles.flipCardTitleButtonSmallRight)}>{item.Title}</div>
                                                                                      </div>
                                                                                      <div className={css(styles.flipCardBack)} style={backstyle}>
                                                                                        <p className={css(styles.flipCardBackTextButton, styles.flipCardTitleButtonSmallRight)}>
                                                                                          {item.BackText}
                                                                                        </p>
                                                                                      </div>
                                                                                    </div>
                                                                                  </div>

                                                                                ) : CardSize == "ButtonSmall" && fonttextalign == "AlignLeft" && fontsizeop == "medium" ?
                                                                                  (
                                                                                    <div className={css(styles.flipCardButtonSmall)}>
                                                                                      <div className={css(styles.flipCardInner)}>
                                                                                        <div className={css(styles.flipCardFront)} style={frontstyle}>

                                                                                          <div className={css(styles.flipCardTitleButtonSmall, styles.flipCardTitleButtonSmallMediumLeft)}>{item.Title}</div>
                                                                                        </div>
                                                                                        <div className={css(styles.flipCardBack)} style={backstyle}>
                                                                                          <p className={css(styles.flipCardBackTextButtonMedium, styles.flipCardTitleCompactMediumLeft)}>
                                                                                            {item.BackText}
                                                                                          </p>
                                                                                        </div>
                                                                                      </div>
                                                                                    </div>

                                                                                  ) : CardSize == "ButtonSmall" && fonttextalign == "AlignCenter" && fontsizeop == "medium" ?
                                                                                    (
                                                                                      <div className={css(styles.flipCardButtonSmall)}>
                                                                                        <div className={css(styles.flipCardInner)}>
                                                                                          <div className={css(styles.flipCardFront)} style={frontstyle}>

                                                                                            <div className={css(styles.flipCardTitleButtonSmall, styles.flipCardTitleButtonSmallMediumCenter)}>{item.Title}</div>
                                                                                          </div>
                                                                                          <div className={css(styles.flipCardBack)} style={backstyle}>
                                                                                            <p className={css(styles.flipCardBackTextButtonMedium, styles.flipCardTitleCompactMediumCenter)}>
                                                                                              {item.BackText}
                                                                                            </p>
                                                                                          </div>
                                                                                        </div>
                                                                                      </div>

                                                                                    ) : CardSize == "ButtonSmall" && fonttextalign == "AlignRight" && fontsizeop == "medium" ?
                                                                                      (
                                                                                        <div className={css(styles.flipCardButtonSmall)}>
                                                                                          <div className={css(styles.flipCardInner)}>
                                                                                            <div className={css(styles.flipCardFront)} style={frontstyle}>

                                                                                              <div className={css(styles.flipCardTitleButtonSmall, styles.flipCardTitleButtonSmallMediumRight)}>{item.Title}</div>
                                                                                            </div>
                                                                                            <div className={css(styles.flipCardBack)} style={backstyle}>
                                                                                              <p className={css(styles.flipCardBackTextButtonMedium, styles.flipCardTitleCompactMediumRight)}>
                                                                                                {item.BackText}
                                                                                              </p>
                                                                                            </div>
                                                                                          </div>
                                                                                        </div>

                                                                                      ) : CardSize == "ButtonSmall" && fonttextalign == "AlignLeft" && fontsizeop == "large" ?
                                                                                        (
                                                                                          <div className={css(styles.flipCardButtonSmall)}>
                                                                                            <div className={css(styles.flipCardInner)}>
                                                                                              <div className={css(styles.flipCardFront)} style={frontstyle}>

                                                                                                <div className={css(styles.flipCardTitleButtonSmall, styles.flipCardTitleButtonSmallLargeLeft)}>{item.Title}</div>
                                                                                              </div>
                                                                                              <div className={css(styles.flipCardBack)} style={backstyle}>
                                                                                                <p className={css(styles.flipCardBackTextButtonLarge, styles.flipCardTitleCompactLargeLeft)}>
                                                                                                  {item.BackText}
                                                                                                </p>
                                                                                              </div>
                                                                                            </div>
                                                                                          </div>

                                                                                        ) : CardSize == "ButtonSmall" && fonttextalign == "AlignCenter" && fontsizeop == "large" ?
                                                                                          (
                                                                                            <div className={css(styles.flipCardButtonSmall)}>
                                                                                              <div className={css(styles.flipCardInner)}>
                                                                                                <div className={css(styles.flipCardFront)} style={frontstyle}>

                                                                                                  <div className={css(styles.flipCardTitleButtonSmall, styles.flipCardTitleButtonSmallLargeCenter)}>{item.Title}</div>
                                                                                                </div>
                                                                                                <div className={css(styles.flipCardBack)} style={backstyle}>
                                                                                                  <p className={css(styles.flipCardBackTextButtonLarge, styles.flipCardTitleButtonSmallLargeCenter)}>
                                                                                                    {item.BackText}
                                                                                                  </p>
                                                                                                </div>
                                                                                              </div>
                                                                                            </div>

                                                                                          ) : CardSize == "ButtonSmall" && fonttextalign == "AlignRight" && fontsizeop == "large" ?
                                                                                            (
                                                                                              <div className={css(styles.flipCardButtonSmall)}>
                                                                                                <div className={css(styles.flipCardInner)}>
                                                                                                  <div className={css(styles.flipCardFront)} style={frontstyle}>

                                                                                                    <div className={css(styles.flipCardTitleButtonSmall, styles.flipCardTitleButtonSmallLargeRight)}>{item.Title}</div>
                                                                                                  </div>
                                                                                                  <div className={css(styles.flipCardBack)} style={backstyle}>
                                                                                                    <p className={css(styles.flipCardBackTextButtonLarge, styles.flipCardTitleButtonSmallLargeRight)}>
                                                                                                      {item.BackText}
                                                                                                    </p>
                                                                                                  </div>
                                                                                                </div>
                                                                                              </div>

                                                                                            ) : CardSize == "ButtonSmall" && fonttextalign == "AlignLeft" && fontsizeop == "small" ?
                                                                                              (
                                                                                                <div className={css(styles.flipCardButtonSmall)}>
                                                                                                  <div className={css(styles.flipCardInner)}>
                                                                                                    <div className={css(styles.flipCardFront)} style={frontstyle}>

                                                                                                      <div className={css(styles.flipCardTitleButtonSmall, styles.flipCardTitleButtonSmallLeft)}>{item.Title}</div>
                                                                                                    </div>
                                                                                                    <div className={css(styles.flipCardBack)} style={backstyle}>
                                                                                                      <p className={css(styles.flipCardBackTextButton, styles.flipCardTitleButtonSmallLeft)}>
                                                                                                        {item.BackText}
                                                                                                      </p>
                                                                                                    </div>
                                                                                                  </div>
                                                                                                </div>

                                                                                              ) : CardSize == "ButtonSmall" && fonttextalign == "AlignCenter" && fontsizeop == "small" ?
                                                                                                (
                                                                                                  <div className={css(styles.flipCardButtonSmall)}>
                                                                                                    <div className={css(styles.flipCardInner)}>
                                                                                                      <div className={css(styles.flipCardFront)} style={frontstyle}>

                                                                                                        <div className={css(styles.flipCardTitleButtonSmall, styles.flipCardTitleButtonSmallCenter)}>{item.Title}</div>
                                                                                                      </div>
                                                                                                      <div className={css(styles.flipCardBack)} style={backstyle}>
                                                                                                        <p className={css(styles.flipCardBackTextButton, styles.flipCardTitleButtonSmallCenter)}>
                                                                                                          {item.BackText}
                                                                                                        </p>
                                                                                                      </div>
                                                                                                    </div>
                                                                                                  </div>

                                                                                                ) : false

                  )

                  :
                  //If FlipCard Not Enabled

                  (
                    CardSize == "Compact" && fonttextalign == "AlignLeft" && fontsizeop == "small" ?
                      (
                        <div className={css(styles.flipCarde)}>
                          <div className={css(styles.flipCardInner)}>
                            <div className={css(styles.flipCardFront)} style={frontstyle}>
                              <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                className={css(styles.flipCardImageSizeCompact)} />
                              <div className={css(styles.flipCardTitleCompactSmall, styles.flipCardTitleCompactSmallLeft)}>{item.Title}</div>
                            </div>
                            <div className={css(styles.flipCardBack)} style={backstyle}>
                              <p className={css(styles.flipCardBackTextCompact, styles.flipCardTitleCompactSmallLeft)}>
                                {item.BackText}
                              </p>
                            </div>
                          </div>
                        </div>

                      ) : CardSize == "Compact" && fonttextalign == "AlignCenter" && fontsizeop == "small" ?
                        (
                          <div className={css(styles.flipCarde)}>
                            <div className={css(styles.flipCardInner)}>
                              <div className={css(styles.flipCardFront)} style={frontstyle}>
                                <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                  className={css(styles.flipCardImageSizeCompact)} />
                                <div className={css(styles.flipCardTitleCompactSmall, styles.flipCardTitleCompactSmallCenter)}>{item.Title}</div>
                              </div>
                              <div className={css(styles.flipCardBack)} style={backstyle}>
                                <p className={css(styles.flipCardBackTextCompact, styles.flipCardTitleCompactSmallCenter)}>
                                  {item.BackText}
                                </p>
                              </div>
                            </div>
                          </div>

                        ) : CardSize == "Compact" && fonttextalign == "AlignRight" && fontsizeop == "small" ?
                          (
                            <div className={css(styles.flipCarde)}>
                              <div className={css(styles.flipCardInner)}>
                                <div className={css(styles.flipCardFront)} style={frontstyle}>
                                  <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                    className={css(styles.flipCardImageSizeCompact)} />
                                  <div className={css(styles.flipCardTitleCompactSmall, styles.flipCardTitleCompactSmallRight)}>{item.Title}</div>
                                </div>
                                <div className={css(styles.flipCardBack)} style={backstyle}>
                                  <p className={css(styles.flipCardBackTextCompact, styles.flipCardTitleCompactSmallRight)}>
                                    {item.BackText}
                                  </p>
                                </div>
                              </div>
                            </div>

                          ) : CardSize == "Compact" && fonttextalign == "AlignLeft" && fontsizeop == "medium" ?
                            (
                              <div className={css(styles.flipCarde)}>
                                <div className={css(styles.flipCardInner)}>
                                  <div className={css(styles.flipCardFront)} style={frontstyle}>
                                    <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                      className={css(styles.flipCardImageSizeCompact)} />
                                    <div className={css(styles.flipCardTitleCompactMedium, styles.flipCardTitleCompactMediumLeft)}>{item.Title}</div>
                                  </div>
                                  <div className={css(styles.flipCardBack)} style={backstyle}>
                                    <p className={css(styles.flipCardBackTextCompact, styles.flipCardTitleCompactMediumLeft)}>
                                      {item.BackText}
                                    </p>
                                  </div>
                                </div>
                              </div>

                            ) : CardSize == "Compact" && fonttextalign == "AlignCenter" && fontsizeop == "medium" ?
                              (
                                <div className={css(styles.flipCarde)}>
                                  <div className={css(styles.flipCardInner)}>
                                    <div className={css(styles.flipCardFront)} style={frontstyle}>
                                      <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                        className={css(styles.flipCardImageSizeCompact)} />
                                      <div className={css(styles.flipCardTitleCompactMedium, styles.flipCardTitleCompactMediumCenter)}>{item.Title}</div>
                                    </div>
                                    <div className={css(styles.flipCardBack)} style={backstyle}>
                                      <p className={css(styles.flipCardBackTextCompact, styles.flipCardTitleCompactMediumCenter)}>
                                        {item.BackText}
                                      </p>
                                    </div>
                                  </div>
                                </div>

                              ) : CardSize == "Compact" && fonttextalign == "AlignRight" && fontsizeop == "medium" ?
                                (
                                  <div className={css(styles.flipCarde)}>
                                    <div className={css(styles.flipCardInner)}>
                                      <div className={css(styles.flipCardFront)} style={frontstyle}>
                                        <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                          className={css(styles.flipCardImageSizeCompact)} />
                                        <div className={css(styles.flipCardTitleCompactMedium, styles.flipCardTitleCompactMediumRight)}>{item.Title}</div>
                                      </div>
                                      <div className={css(styles.flipCardBack)} style={backstyle}>
                                        <p className={css(styles.flipCardBackTextCompact, styles.flipCardTitleCompactMediumRight)}>
                                          {item.BackText}
                                        </p>
                                      </div>
                                    </div>
                                  </div>

                                ) : CardSize == "Compact" && fonttextalign == "AlignLeft" && fontsizeop == "large" ?
                                  (
                                    <div className={css(styles.flipCarde)}>
                                      <div className={css(styles.flipCardInner)}>
                                        <div className={css(styles.flipCardFront)} style={frontstyle}>
                                          <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                            className={css(styles.flipCardImageSizeCompact)} />
                                          <div className={css(styles.flipCardTitleCompactLarge, styles.flipCardTitleCompactLargeLeft)}>{item.Title}</div>
                                        </div>
                                        <div className={css(styles.flipCardBack)} style={backstyle}>
                                          <p className={css(styles.flipCardBackTextCompact, styles.flipCardTitleCompactLargeLeft)}>
                                            {item.BackText}
                                          </p>
                                        </div>
                                      </div>
                                    </div>

                                  ) : CardSize == "Compact" && fonttextalign == "AlignCenter" && fontsizeop == "large" ?
                                    (
                                      <div className={css(styles.flipCarde)}>
                                        <div className={css(styles.flipCardInner)}>
                                          <div className={css(styles.flipCardFront)} style={frontstyle}>
                                            <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                              className={css(styles.flipCardImageSizeCompact)} />
                                            <div className={css(styles.flipCardTitleCompactLarge, styles.flipCardTitleCompactLargeCenter)}>{item.Title}</div>
                                          </div>
                                          <div className={css(styles.flipCardBack)} style={backstyle}>
                                            <p className={css(styles.flipCardBackTextCompact, styles.flipCardTitleCompactLargeCenter)}>
                                              {item.BackText}
                                            </p>
                                          </div>
                                        </div>
                                      </div>

                                    ) : CardSize == "Compact" && fonttextalign == "AlignRight" && fontsizeop == "large" ?
                                      (
                                        <div className={css(styles.flipCarde)}>
                                          <div className={css(styles.flipCardInner)}>
                                            <div className={css(styles.flipCardFront)} style={frontstyle}>
                                              <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                className={css(styles.flipCardImageSizeCompact)} />
                                              <div className={css(styles.flipCardTitleCompactLarge, styles.flipCardTitleCompactLargeRight)}>{item.Title}</div>
                                            </div>
                                            <div className={css(styles.flipCardBack)} style={backstyle}>
                                              <p className={css(styles.flipCardBackTextCompact, styles.flipCardTitleCompactLargeRight)}>
                                                {item.BackText}
                                              </p>
                                            </div>
                                          </div>
                                        </div>

                                      ) : CardSize == "Medium" && fonttextalign == "AlignLeft" && fontsizeop == "small" ?
                                        (
                                          <div className={css(styles.flipCardeMedium)}>
                                            <div className={css(styles.flipCardInner)}>
                                              <div className={css(styles.flipCardFront)} style={frontstyle}>
                                                <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                  className={css(styles.flipCardImageSizeMedium)} />
                                                <div className={css(styles.flipCardTitleMediumSmall, styles.flipCardTitleMediumSmallLeft)}>{item.Title}</div>
                                              </div>
                                              <div className={css(styles.flipCardBack)} style={backstyle}>
                                                <p className={css(styles.flipCardBackTextMedium, styles.flipCardTitleMediumSmallLeft)}>
                                                  {item.BackText}
                                                </p>
                                              </div>
                                            </div>
                                          </div>

                                        ) : CardSize == "Medium" && fonttextalign == "AlignCenter" && fontsizeop == "small" ?
                                          (
                                            <div className={css(styles.flipCardeMedium)}>
                                              <div className={css(styles.flipCardInner)}>
                                                <div className={css(styles.flipCardFront)} style={frontstyle}>
                                                  <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                    className={css(styles.flipCardImageSizeMedium)} />
                                                  <div className={css(styles.flipCardTitleMediumSmall, styles.flipCardTitleMediumSmallCenter)}>{item.Title}</div>
                                                </div>
                                                <div className={css(styles.flipCardBack)} style={backstyle}>
                                                  <p className={css(styles.flipCardBackTextMedium, styles.flipCardTitleMediumSmallCenter)}>
                                                    {item.BackText}
                                                  </p>
                                                </div>
                                              </div>
                                            </div>

                                          ) : CardSize == "Medium" && fonttextalign == "AlignRight" && fontsizeop == "small" ?
                                            (
                                              <div className={css(styles.flipCardeMedium)}>
                                                <div className={css(styles.flipCardInner)}>
                                                  <div className={css(styles.flipCardFront)} style={frontstyle}>
                                                    <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                      className={css(styles.flipCardImageSizeMedium)} />
                                                    <div className={css(styles.flipCardTitleMediumSmall, styles.flipCardTitleMediumSmallRight)}>{item.Title}</div>
                                                  </div>
                                                  <div className={css(styles.flipCardBack)} style={backstyle}>
                                                    <p className={css(styles.flipCardBackTextMedium, styles.flipCardTitleMediumSmallRight)}>
                                                      {item.BackText}
                                                    </p>
                                                  </div>
                                                </div>
                                              </div>

                                            ) : CardSize == "Medium" && fonttextalign == "AlignLeft" && fontsizeop == "medium" ?
                                              (
                                                <div className={css(styles.flipCardeMedium)}>
                                                  <div className={css(styles.flipCardInner)}>
                                                    <div className={css(styles.flipCardFront)} style={frontstyle}>
                                                      <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                        className={css(styles.flipCardImageSizeMedium)} />
                                                      <div className={css(styles.flipCardTitleMediumMedium, styles.flipCardTitleMediumMediumLeft)}>{item.Title}</div>
                                                    </div>
                                                    <div className={css(styles.flipCardBack)} style={backstyle}>
                                                      <p className={css(styles.flipCardBackTextMedium, styles.flipCardTitleMediumMediumLeft)}>
                                                        {item.BackText}
                                                      </p>
                                                    </div>
                                                  </div>
                                                </div>

                                              ) : CardSize == "Medium" && fonttextalign == "AlignCenter" && fontsizeop == "medium" ?
                                                (
                                                  <div className={css(styles.flipCardeMedium)}>
                                                    <div className={css(styles.flipCardInner)}>
                                                      <div className={css(styles.flipCardFront)} style={frontstyle}>
                                                        <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                          className={css(styles.flipCardImageSizeMedium)} />
                                                        <div className={css(styles.flipCardTitleMediumMedium, styles.flipCardTitleMediumMediumCenter)}>{item.Title}</div>
                                                      </div>
                                                      <div className={css(styles.flipCardBack)} style={backstyle}>
                                                        <p className={css(styles.flipCardBackTextMedium, styles.flipCardTitleMediumMediumCenter)}>
                                                          {item.BackText}
                                                        </p>
                                                      </div>
                                                    </div>
                                                  </div>

                                                ) : CardSize == "Medium" && fonttextalign == "AlignRight" && fontsizeop == "medium" ?
                                                  (
                                                    <div className={css(styles.flipCardeMedium)}>
                                                      <div className={css(styles.flipCardInner)}>
                                                        <div className={css(styles.flipCardFront)} style={frontstyle}>
                                                          <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                            className={css(styles.flipCardImageSizeMedium)} />
                                                          <div className={css(styles.flipCardTitleMediumMedium, styles.flipCardTitleMediumMediumRight)}>{item.Title}</div>
                                                        </div>
                                                        <div className={css(styles.flipCardBack)} style={backstyle}>
                                                          <p className={css(styles.flipCardBackTextMedium, styles.flipCardTitleMediumMediumRight)}>
                                                            {item.BackText}
                                                          </p>
                                                        </div>
                                                      </div>
                                                    </div>

                                                  ) : CardSize == "Medium" && fonttextalign == "AlignLeft" && fontsizeop == "large" ?
                                                    (
                                                      <div className={css(styles.flipCardeMedium)}>
                                                        <div className={css(styles.flipCardInner)}>
                                                          <div className={css(styles.flipCardFront)} style={frontstyle}>
                                                            <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                              className={css(styles.flipCardImageSizeMedium)} />
                                                            <div className={css(styles.flipCardTitleMediumMedium, styles.flipCardTitleMediumLargeLeft)}>{item.Title}</div>
                                                          </div>
                                                          <div className={css(styles.flipCardBack)} style={backstyle}>
                                                            <p className={css(styles.flipCardBackTextMedium, styles.flipCardTitleMediumLargeLeft)}>
                                                              {item.BackText}
                                                            </p>
                                                          </div>
                                                        </div>
                                                      </div>

                                                    ) : CardSize == "Medium" && fonttextalign == "AlignCenter" && fontsizeop == "large" ?
                                                      (
                                                        <div className={css(styles.flipCardeMedium)}>
                                                          <div className={css(styles.flipCardInner)}>
                                                            <div className={css(styles.flipCardFront)} style={frontstyle}>
                                                              <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                                className={css(styles.flipCardImageSizeMedium)} />
                                                              <div className={css(styles.flipCardTitleMediumMedium, styles.flipCardTitleMediumLargeCenter)}>{item.Title}</div>
                                                            </div>
                                                            <div className={css(styles.flipCardBack)} style={backstyle}>
                                                              <p className={css(styles.flipCardBackTextMedium, styles.flipCardTitleMediumLargeCenter)}>
                                                                {item.BackText}
                                                              </p>
                                                            </div>
                                                          </div>
                                                        </div>

                                                      ) : CardSize == "Medium" && fonttextalign == "AlignRight" && fontsizeop == "large" ?
                                                        (
                                                          <div className={css(styles.flipCardeMedium)}>
                                                            <div className={css(styles.flipCardInner)}>
                                                              <div className={css(styles.flipCardFront)} style={frontstyle}>
                                                                <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                                  className={css(styles.flipCardImageSizeMedium)} />
                                                                <div className={css(styles.flipCardTitleMediumMedium, styles.flipCardTitleMediumLargeRight)}>{item.Title}</div>
                                                              </div>
                                                              <div className={css(styles.flipCardBack)} style={backstyle}>
                                                                <p className={css(styles.flipCardBackTextMedium, styles.flipCardTitleMediumLargeRight)}>
                                                                  {item.BackText}
                                                                </p>
                                                              </div>
                                                            </div>
                                                          </div>

                                                        ) : CardSize == "Large" && fonttextalign == "AlignLeft" && fontsizeop == "small" ?
                                                          (
                                                            <div className={css(styles.flipCardeLarge)}>
                                                              <div className={css(styles.flipCardInner)}>
                                                                <div className={css(styles.flipCardFront)} style={frontstyle}>
                                                                  <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                                    className={css(styles.flipCardImageSizeLarge)} />
                                                                  <div className={css(styles.flipCardTitleLargeSmall, styles.flipCardTitleLargeSmallLeft)}>{item.Title}</div>
                                                                </div>
                                                                <div className={css(styles.flipCardBack)} style={backstyle}>
                                                                  <p className={css(styles.flipCardBackTextLarge, styles.flipCardTitleLargeSmallLeft)}>
                                                                    {item.BackText}
                                                                  </p>
                                                                </div>
                                                              </div>
                                                            </div>

                                                          ) : CardSize == "Large" && fonttextalign == "AlignCenter" && fontsizeop == "small" ?
                                                            (
                                                              <div className={css(styles.flipCardeLarge)}>
                                                                <div className={css(styles.flipCardInner)}>
                                                                  <div className={css(styles.flipCardFront)} style={frontstyle}>
                                                                    <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                                      className={css(styles.flipCardImageSizeLarge)} />
                                                                    <div className={css(styles.flipCardTitleLargeSmall, styles.flipCardTitleLargeSmallCenter)}>{item.Title}</div>
                                                                  </div>
                                                                  <div className={css(styles.flipCardBack)} style={backstyle}>
                                                                    <p className={css(styles.flipCardBackTextLarge, styles.flipCardTitleLargeSmallCenter)}>
                                                                      {item.BackText}
                                                                    </p>
                                                                  </div>
                                                                </div>
                                                              </div>

                                                            ) : CardSize == "Large" && fonttextalign == "AlignRight" && fontsizeop == "small" ?
                                                              (
                                                                <div className={css(styles.flipCardeLarge)}>
                                                                  <div className={css(styles.flipCardInner)}>
                                                                    <div className={css(styles.flipCardFront)} style={frontstyle}>
                                                                      <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                                        className={css(styles.flipCardImageSizeLarge)} />
                                                                      <div className={css(styles.flipCardTitleLargeSmall, styles.flipCardTitleLargeSmallRight)}>{item.Title}</div>
                                                                    </div>
                                                                    <div className={css(styles.flipCardBack)} style={backstyle}>
                                                                      <p className={css(styles.flipCardBackTextLarge, styles.flipCardTitleLargeSmallRight)}>
                                                                        {item.BackText}
                                                                      </p>
                                                                    </div>
                                                                  </div>
                                                                </div>

                                                              ) : CardSize == "Large" && fonttextalign == "AlignLeft" && fontsizeop == "medium" ?
                                                                (
                                                                  <div className={css(styles.flipCardeLarge)}>
                                                                    <div className={css(styles.flipCardInner)}>
                                                                      <div className={css(styles.flipCardFront)} style={frontstyle}>
                                                                        <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                                          className={css(styles.flipCardImageSizeLarge)} />
                                                                        <div className={css(styles.flipCardTitleLargeMedium, styles.flipCardTitleLargeMediumLeft)}>{item.Title}</div>
                                                                      </div>
                                                                      <div className={css(styles.flipCardBack)} style={backstyle}>
                                                                        <p className={css(styles.flipCardBackTextLarge, styles.flipCardTitleLargeMediumLeft)}>
                                                                          {item.BackText}
                                                                        </p>
                                                                      </div>
                                                                    </div>
                                                                  </div>

                                                                ) : CardSize == "Large" && fonttextalign == "AlignCenter" && fontsizeop == "medium" ?
                                                                  (
                                                                    <div className={css(styles.flipCardeLarge)}>
                                                                      <div className={css(styles.flipCardInner)}>
                                                                        <div className={css(styles.flipCardFront)} style={frontstyle}>
                                                                          <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                                            className={css(styles.flipCardImageSizeLarge)} />
                                                                          <div className={css(styles.flipCardTitleLargeMedium, styles.flipCardTitleLargeMediumCenter)}>{item.Title}</div>
                                                                        </div>
                                                                        <div className={css(styles.flipCardBack)} style={backstyle}>
                                                                          <p className={css(styles.flipCardBackTextLarge, styles.flipCardTitleLargeMediumCenter)}>
                                                                            {item.BackText}
                                                                          </p>
                                                                        </div>
                                                                      </div>
                                                                    </div>

                                                                  ) : CardSize == "Large" && fonttextalign == "AlignRight" && fontsizeop == "medium" ?
                                                                    (
                                                                      <div className={css(styles.flipCardeLarge)}>
                                                                        <div className={css(styles.flipCardInner)}>
                                                                          <div className={css(styles.flipCardFront)} style={frontstyle}>
                                                                            <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                                              className={css(styles.flipCardImageSizeLarge)} />
                                                                            <div className={css(styles.flipCardTitleLargeMedium, styles.flipCardTitleLargeMediumRight)}>{item.Title}</div>
                                                                          </div>
                                                                          <div className={css(styles.flipCardBack)} style={backstyle}>
                                                                            <p className={css(styles.flipCardBackTextLarge, styles.flipCardTitleLargeMediumRight)}>
                                                                              {item.BackText}
                                                                            </p>
                                                                          </div>
                                                                        </div>
                                                                      </div>

                                                                    ) : CardSize == "Large" && fonttextalign == "AlignLeft" && fontsizeop == "large" ?
                                                                      (
                                                                        <div className={css(styles.flipCardeLarge)}>
                                                                          <div className={css(styles.flipCardInner)}>
                                                                            <div className={css(styles.flipCardFront)} style={frontstyle}>
                                                                              <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                                                className={css(styles.flipCardImageSizeLarge)} />
                                                                              <div className={css(styles.flipCardTitleLargeLarge, styles.flipCardTitleLargeLargeLeft)}>{item.Title}</div>
                                                                            </div>
                                                                            <div className={css(styles.flipCardBack)} style={backstyle}>
                                                                              <p className={css(styles.flipCardBackTextLarge, styles.flipCardTitleLargeLargeLeft)}>
                                                                                {item.BackText}
                                                                              </p>
                                                                            </div>
                                                                          </div>
                                                                        </div>

                                                                      ) : CardSize == "Large" && fonttextalign == "AlignCenter" && fontsizeop == "large" ?
                                                                        (
                                                                          <div className={css(styles.flipCardeLarge)}>
                                                                            <div className={css(styles.flipCardInner)}>
                                                                              <div className={css(styles.flipCardFront)} style={frontstyle}>
                                                                                <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                                                  className={css(styles.flipCardImageSizeLarge)} />
                                                                                <div className={css(styles.flipCardTitleLargeLarge, styles.flipCardTitleLargeLargeCenter)}>{item.Title}</div>
                                                                              </div>
                                                                              <div className={css(styles.flipCardBack)} style={backstyle}>
                                                                                <p className={css(styles.flipCardBackTextLarge, styles.flipCardTitleLargeLargeCenter)}>
                                                                                  {item.BackText}
                                                                                </p>
                                                                              </div>
                                                                            </div>
                                                                          </div>

                                                                        ) : CardSize == "Large" && fonttextalign == "AlignRight" && fontsizeop == "large" ?
                                                                          (
                                                                            <div className={css(styles.flipCardeLarge)}>
                                                                              <div className={css(styles.flipCardInner)}>
                                                                                <div className={css(styles.flipCardFront)} style={frontstyle}>
                                                                                  <img src={item.filePicker.fileAbsoluteUrl != null ? item.filePicker.fileAbsoluteUrl : item.filePicker.ServerRelativeUrl}
                                                                                    className={css(styles.flipCardImageSizeLarge)} />
                                                                                  <div className={css(styles.flipCardTitleLargeLarge, styles.flipCardTitleLargeLargeRight)}>{item.Title}</div>
                                                                                </div>
                                                                                <div className={css(styles.flipCardBack)} style={backstyle}>
                                                                                  <p className={css(styles.flipCardBackTextLarge, styles.flipCardTitleLargeLargeRight)}>
                                                                                    {item.BackText}
                                                                                  </p>
                                                                                </div>
                                                                              </div>
                                                                            </div>

                                                                          ) : CardSize == "ButtonSmall" && fonttextalign == "AlignLeft" && fontsizeop == "small" ?
                                                                            (
                                                                              <div className={css(styles.flipCardButtonSmall)}>
                                                                                <div className={css(styles.flipCardInner)}>
                                                                                  <div className={css(styles.flipCardFront)} style={frontstyle}>

                                                                                    <div className={css(styles.flipCardTitleButtonSmall, styles.flipCardTitleButtonSmallLeft)}>{item.Title}</div>
                                                                                  </div>
                                                                                  <div className={css(styles.flipCardBack)} style={backstyle}>
                                                                                    <p className={css(styles.flipCardBackTextButton, styles.flipCardTitleButtonSmallLeft)}>
                                                                                      {item.BackText}
                                                                                    </p>
                                                                                  </div>
                                                                                </div>
                                                                              </div>

                                                                            ) : CardSize == "ButtonSmall" && fonttextalign == "AlignCenter" && fontsizeop == "small" ?
                                                                              (
                                                                                <div className={css(styles.flipCardButtonSmall)}>
                                                                                  <div className={css(styles.flipCardInner)}>
                                                                                    <div className={css(styles.flipCardFront)} style={frontstyle}>

                                                                                      <div className={css(styles.flipCardTitleButtonSmall, styles.flipCardTitleButtonSmallCenter)}>{item.Title}</div>
                                                                                    </div>
                                                                                    <div className={css(styles.flipCardBack)} style={backstyle}>
                                                                                      <p className={css(styles.flipCardBackTextButton, styles.flipCardTitleButtonSmallCenter)}>
                                                                                        {item.BackText}
                                                                                      </p>
                                                                                    </div>
                                                                                  </div>
                                                                                </div>

                                                                              ) : CardSize == "ButtonSmall" && fonttextalign == "AlignRight" && fontsizeop == "small" ?
                                                                                (
                                                                                  <div className={css(styles.flipCardButtonSmall)}>
                                                                                    <div className={css(styles.flipCardInner)}>
                                                                                      <div className={css(styles.flipCardFront)} style={frontstyle}>

                                                                                        <div className={css(styles.flipCardTitleButtonSmall, styles.flipCardTitleButtonSmallRight)}>{item.Title}</div>
                                                                                      </div>
                                                                                      <div className={css(styles.flipCardBack)} style={backstyle}>
                                                                                        <p className={css(styles.flipCardBackTextButton, styles.flipCardTitleButtonSmallRight)}>
                                                                                          {item.BackText}
                                                                                        </p>
                                                                                      </div>
                                                                                    </div>
                                                                                  </div>

                                                                                ) : CardSize == "ButtonSmall" && fonttextalign == "AlignLeft" && fontsizeop == "medium" ?
                                                                                  (
                                                                                    <div className={css(styles.flipCardButtonSmall)}>
                                                                                      <div className={css(styles.flipCardInner)}>
                                                                                        <div className={css(styles.flipCardFront)} style={frontstyle}>

                                                                                          <div className={css(styles.flipCardTitleButtonSmall, styles.flipCardTitleButtonSmallMediumLeft)}>{item.Title}</div>
                                                                                        </div>
                                                                                        <div className={css(styles.flipCardBack)} style={backstyle}>
                                                                                          <p className={css(styles.flipCardBackTextButtonMedium, styles.flipCardTitleCompactMediumLeft)}>
                                                                                            {item.BackText}
                                                                                          </p>
                                                                                        </div>
                                                                                      </div>
                                                                                    </div>

                                                                                  ) : CardSize == "ButtonSmall" && fonttextalign == "AlignCenter" && fontsizeop == "medium" ?
                                                                                    (
                                                                                      <div className={css(styles.flipCardButtonSmall)}>
                                                                                        <div className={css(styles.flipCardInner)}>
                                                                                          <div className={css(styles.flipCardFront)} style={frontstyle}>

                                                                                            <div className={css(styles.flipCardTitleButtonSmall, styles.flipCardTitleButtonSmallMediumCenter)}>{item.Title}</div>
                                                                                          </div>
                                                                                          <div className={css(styles.flipCardBack)} style={backstyle}>
                                                                                            <p className={css(styles.flipCardBackTextButtonMedium, styles.flipCardTitleCompactMediumCenter)}>
                                                                                              {item.BackText}
                                                                                            </p>
                                                                                          </div>
                                                                                        </div>
                                                                                      </div>

                                                                                    ) : CardSize == "ButtonSmall" && fonttextalign == "AlignRight" && fontsizeop == "medium" ?
                                                                                      (
                                                                                        <div className={css(styles.flipCardButtonSmall)}>
                                                                                          <div className={css(styles.flipCardInner)}>
                                                                                            <div className={css(styles.flipCardFront)} style={frontstyle}>

                                                                                              <div className={css(styles.flipCardTitleButtonSmall, styles.flipCardTitleButtonSmallMediumRight)}>{item.Title}</div>
                                                                                            </div>
                                                                                            <div className={css(styles.flipCardBack)} style={backstyle}>
                                                                                              <p className={css(styles.flipCardBackTextButtonMedium, styles.flipCardTitleCompactMediumRight)}>
                                                                                                {item.BackText}
                                                                                              </p>
                                                                                            </div>
                                                                                          </div>
                                                                                        </div>

                                                                                      ) : CardSize == "ButtonSmall" && fonttextalign == "AlignLeft" && fontsizeop == "large" ?
                                                                                        (
                                                                                          <div className={css(styles.flipCardButtonSmall)}>
                                                                                            <div className={css(styles.flipCardInner)}>
                                                                                              <div className={css(styles.flipCardFront)} style={frontstyle}>

                                                                                                <div className={css(styles.flipCardTitleButtonSmall, styles.flipCardTitleButtonSmallLargeLeft)}>{item.Title}</div>
                                                                                              </div>
                                                                                              <div className={css(styles.flipCardBack)} style={backstyle}>
                                                                                                <p className={css(styles.flipCardBackTextButtonLarge, styles.flipCardTitleCompactLargeLeft)}>
                                                                                                  {item.BackText}
                                                                                                </p>
                                                                                              </div>
                                                                                            </div>
                                                                                          </div>

                                                                                        ) : CardSize == "ButtonSmall" && fonttextalign == "AlignCenter" && fontsizeop == "large" ?
                                                                                          (
                                                                                            <div className={css(styles.flipCardButtonSmall)}>
                                                                                              <div className={css(styles.flipCardInner)}>
                                                                                                <div className={css(styles.flipCardFront)} style={frontstyle}>

                                                                                                  <div className={css(styles.flipCardTitleButtonSmall, styles.flipCardTitleButtonSmallLargeCenter)}>{item.Title}</div>
                                                                                                </div>
                                                                                                <div className={css(styles.flipCardBack)} style={backstyle}>
                                                                                                  <p className={css(styles.flipCardBackTextButtonLarge, styles.flipCardTitleButtonSmallLargeCenter)}>
                                                                                                    {item.BackText}
                                                                                                  </p>
                                                                                                </div>
                                                                                              </div>
                                                                                            </div>

                                                                                          ) : CardSize == "ButtonSmall" && fonttextalign == "AlignRight" && fontsizeop == "large" ?
                                                                                            (
                                                                                              <div className={css(styles.flipCardButtonSmall)}>
                                                                                                <div className={css(styles.flipCardInner)}>
                                                                                                  <div className={css(styles.flipCardFront)} style={frontstyle}>

                                                                                                    <div className={css(styles.flipCardTitleButtonSmall, styles.flipCardTitleButtonSmallLargeRight)}>{item.Title}</div>
                                                                                                  </div>
                                                                                                  <div className={css(styles.flipCardBack)} style={backstyle}>
                                                                                                    <p className={css(styles.flipCardBackTextButtonLarge, styles.flipCardTitleButtonSmallLargeRight)}>
                                                                                                      {item.BackText}
                                                                                                    </p>
                                                                                                  </div>
                                                                                                </div>
                                                                                              </div>

                                                                                            ) : CardSize == "ButtonSmall" && fonttextalign == "AlignLeft" && fontsizeop == "small" ?
                                                                                              (
                                                                                                <div className={css(styles.flipCardButtonSmall)}>
                                                                                                  <div className={css(styles.flipCardInner)}>
                                                                                                    <div className={css(styles.flipCardFront)} style={frontstyle}>

                                                                                                      <div className={css(styles.flipCardTitleButtonSmall, styles.flipCardTitleButtonSmallLeft)}>{item.Title}</div>
                                                                                                    </div>
                                                                                                    <div className={css(styles.flipCardBack)} style={backstyle}>
                                                                                                      <p className={css(styles.flipCardBackTextButton, styles.flipCardTitleButtonSmallLeft)}>
                                                                                                        {item.BackText}
                                                                                                      </p>
                                                                                                    </div>
                                                                                                  </div>
                                                                                                </div>

                                                                                              ) : CardSize == "ButtonSmall" && fonttextalign == "AlignCenter" && fontsizeop == "small" ?
                                                                                                (
                                                                                                  <div className={css(styles.flipCardButtonSmall)}>
                                                                                                    <div className={css(styles.flipCardInner)}>
                                                                                                      <div className={css(styles.flipCardFront)} style={frontstyle}>

                                                                                                        <div className={css(styles.flipCardTitleButtonSmall, styles.flipCardTitleButtonSmallCenter)}>{item.Title}</div>
                                                                                                      </div>
                                                                                                      <div className={css(styles.flipCardBack)} style={backstyle}>
                                                                                                        <p className={css(styles.flipCardBackTextButton, styles.flipCardTitleButtonSmallCenter)}>
                                                                                                          {item.BackText}
                                                                                                        </p>
                                                                                                      </div>
                                                                                                    </div>
                                                                                                  </div>

                                                                                                ) : false

                  )

            }

          </div >



        );
      });

    }


    return (
       <div className={`${styles.imageGallery} ${alignmentClass}`}>
        <div className={styles.container} dir="ltr">

          <div className={css(styles.row)}>
            <div className={css(styles.column, styles.mslg12)}>
              <div className={styles.panelBody}>

                {
                  this.state.showLoader
                    ? <Spinner size={SpinnerSize.large} label="loading..." className={css(styles.loader)} />
                    : ""
                }

                <div className={css(styles.row, styles.mainContent)}>

                  {result.length > 0 ? result : ""}
                  {!result.length && this.state.itemsNotFound ? <MessageBar
                    messageBarType={MessageBarType.warning}
                    isMultiline={false}
                    // onDismiss={log('test')}
                    dismissButtonAriaLabel="Close"
                  >
                    Items not found. Try different search keyword
                  </MessageBar> : ""}
                  <Panel
                    isOpen={this.state.showPanel}
                    type={PanelType.medium}
                    onDismiss={() => this.setState({ showPanel: false })}
                    headerText={this.state.selectedImage.Title}
                  >
                    <div className={styles.modalContent}>
                      <div className={styles.modalBody}>
                        <div className={styles.thumbnail}>
                          <img src={this.state.selectedImage.LinkFilename} title={this.state.selectedImage.Title} id={this.state.selectedImage.Id} />
                        </div>
                        <h3>Tags</h3>
                        {this.state.selectedImage.Department ?
                          <ul className={styles.listGroup}>
                            {
                              tagList
                            }
                          </ul> : ""}
                      </div>
                    </div>
                  </Panel>

                </div>
              </div>
            </div>
          </div>

        </div>
      </div>
    );
  }

}