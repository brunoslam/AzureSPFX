import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'SarWebPartStrings';
import Sar from './components/Sar';
import { ISarProps } from './components/ISarProps';
import { HttpClient, SPHttpClient, HttpClientConfiguration, HttpClientResponse, ODataVersion, IHttpClientConfiguration, IHttpClientOptions, ISPHttpClientOptions } from '@microsoft/sp-http';
export interface ISarWebPartProps {
  description: string;
  context_: any

}
import * as $ from 'jquery';
import Axios from 'axios';

export default class SarWebPart extends BaseClientSideWebPart<ISarWebPartProps> {
  constructor(){
    super()
  }
  public render(): void {
    const element: React.ReactElement<ISarProps > = React.createElement(
      Sar,
      {
        description: this.properties.description,
        context_: this.context
      }
    );

    ReactDom.render(element, this.domElement);

    
   



    /*
    this.context.httpClient
    .get('https://expressindicadoresls.azurewebsites.net/?callback=?', HttpClient.configurations.v1, postOptions)
    .then((res: HttpClientResponse): Promise<any> => {
      return res.json();
    }, (err: any): void => {
      console.log(err);
    }).then((data: any): void => {
      if(data != null){
        data = JSON.parse(data);
      }
    }, (err: any): void => {
      console.log(err);
    });*/
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
