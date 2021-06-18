import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'FindAndReplaceWebPartStrings';
import FindAndReplace from './components/FindAndReplace';
import { IFindAndReplaceProps } from './components/IFindAndReplaceProps';
import { sp } from "@pnp/sp/presets/all";

const context: any = {};
export const FindAndReplaceContext: any = React.createContext(context);

export interface IFindAndReplaceWebPartProps {
  description: string;
  SiteName: string;
 
}

export default class FindAndReplaceWebPart extends BaseClientSideWebPart<IFindAndReplaceWebPartProps> {

  public render(): void {
    
    // sp.setup({
    //   spfxContext: this.context
    // });

    const context : any = this.context;
    const siteURL : string = context._pageContext.web.absoluteUrl.split(context._pageContext.web.serverRelativeUrl)[0];
    console.log(siteURL);

    const element: React.ReactElement<IFindAndReplaceProps> = React.createElement(
      FindAndReplace,
      {
        description: this.properties.description,
        context: this.context,
        SiteName: siteURL + '/sites/' + this.properties.SiteName,
        
      }
    );
    ReactDom.render(element, this.domElement);
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
              groupName: "Find And Replace Webpart",
              groupFields: [
                
                PropertyPaneTextField('SiteName', {
                  label: 'Site Name',
                }),

              ]
            }
          ]
        }
      ]
    };
  }
}
