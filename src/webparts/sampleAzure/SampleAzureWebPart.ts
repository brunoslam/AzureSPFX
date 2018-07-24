import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import axios, { AxiosRequestConfig, AxiosPromise } from 'axios';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './SampleAzureWebPart.module.scss';
import * as strings from 'SampleAzureWebPartStrings';

export interface ISampleAzureWebPartProps {
  description: string;
}




export default class SampleAzureWebPart extends BaseClientSideWebPart<ISampleAzureWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.sampleAzure }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">Welcome to SharePoint!</span>
              <p class="${ styles.subTitle }">Customize SharePoint experiences using Web Parts.</p>
              <p class="${ styles.description }">${escape(this.properties.description)}</p>
              <a href="https://aka.ms/spfx" class="${ styles.button }">
                <span class="${ styles.label }">Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>`;
      
      const tablestorageUrl =  'https://storagelatintest.table.core.windows.net/Persona?st=2018-07-23T22%3A51%3A35Z&se=2018-07-24T22%3A51%3A35Z&sp=r&sv=2018-03-28&tn=persona&sig=b5U%2Bt%2FmqsifokW6sZTonnAP4RNCyHlkuaJZl1DmRIKc%3D';
      axios.get(tablestorageUrl, {
        headers: {
          Accepts: 'application/json'
        }}).then((response) => {
          console.log(response.data);
      }).catch((err)=>{
          console.log(err);

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
