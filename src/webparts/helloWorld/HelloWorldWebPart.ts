import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import styles from './HelloWorldWebPart.module.scss';
import * as strings from 'HelloWorldWebPartStrings';

import {
  ICustomGraphService,
  CustomGraphService,
  ICustomService,
  CustomService,
  ICustomSPService,
  CustomSPService
} from '../../services';

export interface IHelloWorldWebPartProps {
  description: string;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {

  private _customGraphServiceInstance: ICustomGraphService;
  private _customServiceInstance: ICustomService;
  private _customSPServiceInstance: ICustomSPService;

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.helloWorld}">
        <div class="${ styles.container}">
            <div class="${styles.row}">
              <div class="${ styles.column}">
                <span class="${styles.title}">Consuming MSGraphClient, AadHttpClient and SPHttpClient 
                through custom services without passing in SPFx component context.</span>
              </div>
            </div>
            <div class="${styles.row}">
              <span class="${styles.title}">MSGraphClient result:</span>
              <div id="graphResultContainer"></div>
            </div>
            <div class="${styles.row}">
              <span class="${styles.title}">AadHttpClient result:</span>
              <div id="aadClientResultContainer"></div>
            </div>
            <div class="${styles.row}">
              <span class="${styles.title}">SPHttpClient result:</span>
              <div id="spClientResultContainer"></div>
            </div>
          </div>
      </div>`;

    //MSGraphClient
    this._customGraphServiceInstance = this.context.serviceScope.consume(CustomGraphService.serviceKey);
    this._customGraphServiceInstance.getMyDetails().then((user: JSON) => {

      document.getElementById("graphResultContainer").innerText = JSON.stringify(user);

    });

    //AadHttpClient
    this._customServiceInstance = this.context.serviceScope.consume(CustomService.serviceKey);
    this._customServiceInstance.executeMyRequest().then((user: JSON) => {

      document.getElementById("aadClientResultContainer").innerText = JSON.stringify(user);

    });

    //SPHttpClient
    this._customSPServiceInstance = this.context.serviceScope.consume(CustomSPService.serviceKey);
    this._customSPServiceInstance.getWebDetails().then((web: JSON) => {

      document.getElementById("spClientResultContainer").innerText = JSON.stringify(web);

    });
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
