import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import { sp } from '@pnp/sp/presets/all';
import * as React from 'react';
import * as ReactDom from 'react-dom';
import * as strings from 'SaccWebPartStrings';
import { ISaccProps } from './components/ISaccProps';
import Sacc from './components/Sacc';


export interface ISaccWebPartProps {
  description: string;
  context: WebPartContext;
}

export default class SaccWebPart extends BaseClientSideWebPart<ISaccWebPartProps> {
  public onInit(): Promise<void> {

    return super.onInit().then(_ => {

      // other init code may be present

      sp.setup({
        spfxContext: this.context
      });
    });


  }

  public async render(): Promise<void> {
    await import('../sacc/components/customWorkbenchStyles.module.scss');


    const element: React.ReactElement<ISaccProps> = React.createElement(
      Sacc,
      {
        description: this.properties.description,
        context: this.context
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
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
