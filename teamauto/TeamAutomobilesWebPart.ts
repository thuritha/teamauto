import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'TeamAutomobilesWebPartStrings';
import TeamAutomobiles from './components/TeamAutomobiles';
import { ITeamAutomobilesProps } from './components/ITeamAutomobilesProps';
//import { spOperation } from './Services/spService';
import { sp } from '@pnp/sp/presets/all';


export interface ITeamAutomobilesWebPartProps {
  description: string;
}

export default class TeamAutomobilesWebPart extends BaseClientSideWebPart<ITeamAutomobilesWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ITeamAutomobilesProps> = React.createElement(
      TeamAutomobiles,
      {
        description: this.properties.description,
        context : this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }
  protected onInit():Promise<void>{
    console.log("onInit CAlled");
    return super.onInit().then((_) => {
      sp.setup ({
        spfxContext: this.context
      });
    });
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  /*protected get dataVersion(): Version {
    return Version.parse('1.0');
  }*/

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
