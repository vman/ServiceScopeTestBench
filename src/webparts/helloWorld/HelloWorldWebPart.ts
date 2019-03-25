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
            <div class="${ styles.row}">
              <div class="${ styles.column}">
                <span class="${styles.label}">Consume MSGraphClient, AadHttpClient and SPHttpClient 
                thorough custom services without passing in SPFx component context.</span>
              </div>
            </div>
          </div>
      </div>`;

    //MSGraphClient
    this._customGraphServiceInstance = this.context.serviceScope.consume(CustomGraphService.serviceKey);
    this._customGraphServiceInstance.getMyDetails().then((user: JSON)=>{
      console.log(user);
    });

    //AadHttpClient
    this._customServiceInstance = this.context.serviceScope.consume(CustomService.serviceKey);
    this._customServiceInstance.executeMyRequest().then((user: JSON)=>{
      console.log(user);
    });

    //SPHttpClient
    this._customSPServiceInstance = this.context.serviceScope.consume(CustomSPService.serviceKey);
    this._customSPServiceInstance.getWebDetails().then((web: JSON)=>{
      console.log(web);
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
