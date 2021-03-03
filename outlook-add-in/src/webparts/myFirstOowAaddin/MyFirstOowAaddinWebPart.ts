import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './MyFirstOowAaddinWebPart.module.scss';
import * as strings from 'MyFirstOowAaddinWebPartStrings';

export interface IMyFirstOowAaddinWebPartProps {
  description: string;
}

export default class MyFirstOowAaddinWebPart extends BaseClientSideWebPart<IMyFirstOowAaddinWebPartProps> {

  public render(): void {

    let title: string = '';
    let subTitle: string = '';
    let contextDetail: string = '';

    if (this.context.sdks.office) {
      // We have Office context for the solution
      title = "Welcome to Office!";
      subTitle = "Extending Office with custom business extensions.";
      contextDetail = "We are in the context of following email: " + this.context.sdks.office.context.mailbox.userProfile.emailAddress;
    }
    else {
      // We are rendered in normal SharePoint context
      title = "Welcome to SharePoint!";
      subTitle = "Customize SharePoint experiences using Web Parts.";
      contextDetail = "We are in the context of following site: " + this.context.pageContext.web.title;
    }

    this.domElement.innerHTML = `
    <div class="${styles.myFirstOwAaddin}">
        <div class="${styles.container}">
        <div class="${styles.row}">
            <div class="${styles.column}">
            <span class="${styles.title}">${title}</span>
            <p class="${styles.subTitle}">${subTitle}</p>
            <p class="${styles.description}">${contextDetail}</p>
            <p class="${styles.description}">${escape(this.properties.description)}</p>
            <a href="https://aka.ms/spfx" class="${styles.button}">
                <span class="${styles.label}">Learn more</span>
            </a>
            </div>
        </div>
        </div>
    </div>`;

  }

  //@ts-ignore
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
