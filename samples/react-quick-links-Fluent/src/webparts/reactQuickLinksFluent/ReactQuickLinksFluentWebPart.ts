import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'ReactQuickLinksFluentWebPartStrings';
import ReactQuickLinksFluent from './components/ReactQuickLinksFluent';
import { IReactQuickLinksFluentProps } from './components/IReactQuickLinksFluentProps';

export interface IReactQuickLinksFluentWebPartProps {
  description: string;
  listName:string;
  groupBy:string;
  quickLinkColor:string;
  quickLinkColor2:string;
  fontIconColor:string;
  margin:number;
  padding:number;
  maxWidth:number;
  minHeight:number;
  gridWidth:number;
 

}

export default class ReactQuickLinksFluentWebPart extends BaseClientSideWebPart<IReactQuickLinksFluentWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IReactQuickLinksFluentProps> = React.createElement(
      ReactQuickLinksFluent,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
         context: this.context,
        listName: this.properties.listName,
         groupBy: this.properties.groupBy,
       quickLinkColor:this.properties.quickLinkColor,
       quickLinkColor2:this.properties.quickLinkColor2,
       fontIconColor:this.properties.fontIconColor,
       margin: this.properties.margin,
       padding: this.properties.padding,
       maxWidth: this.properties.maxWidth,
       minHeight: this.properties.minHeight,
       gridWidth: this.properties.gridWidth
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
             
              groupFields: [
       
               
                PropertyPaneTextField('listName', {
                  label: "List Name"
                }),
                PropertyPaneTextField('groupBy', {
                  label: "Group by? eg: GROUP eq 'GROUP1'"
                }),
              //   PropertyPaneTextField('quickLinkColor', {
              //    label: "Color for Quick Link"
              //  }),  PropertyPaneTextField('quickLinkColor2', {
             //     label: "Color for Quick Link 2"
              //  }), PropertyPaneTextField('fontIconColor', {
             //     label: "Color for Font & Icon"
              //  }), 
                PropertyPaneSlider('margin', {
                  label: "Margin (default:0)",
                  min: -5,
                  max: 10
                }),PropertyPaneSlider('padding', {
                  label: "Padding (default:20)",
                  min: 0,
                  max: 50
                }),PropertyPaneSlider('maxWidth', {
                  label: "Max Width (default:70)",
                  min: 0,
                  max: 200
                }),PropertyPaneSlider('minHeight', {
                  label: "Min Height (default:50)",
                  min: 0,
                  max: 150
                }),PropertyPaneSlider('gridWidth', {
                  label: "Grid Width (default:600)",
                  min: 200,
                  max: 1500
                })
         
              ]
            }
          ]
        }
      ]
    };
  }
}
