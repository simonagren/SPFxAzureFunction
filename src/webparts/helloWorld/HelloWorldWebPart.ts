import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { AadHttpClient, HttpClientResponse, IHttpClientOptions } from '@microsoft/sp-http';

import * as strings from 'HelloWorldWebPartStrings';
import HelloWorld from './components/HelloWorld';
import { IHelloWorldProps } from './components/IHelloWorldProps';

export interface IHelloWorldWebPartProps {
  description: string;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {
  private client: AadHttpClient;

  public onInit(): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      this.context.aadHttpClientFactory
        // The id of you Azure Function
        .getClient('<ApplicationId>')
        .then((client: AadHttpClient): void => {
          this.client = client;
          resolve();
        }, err => reject(err));
    });
  }

  public async render(): Promise<void> {
    // Setup the options with header and body
    const headers: Headers = new Headers();
    headers.append("Content-type", "application/json");

    const postOptions: IHttpClientOptions = {
      headers: headers,
      body: '{ "site": "<SiteName>"}'
    };


    const result = await this.client
      .post('https://<FunctionApp>.azurewebsites.net/api/PnPNewInvite', AadHttpClient.configurations.v1, postOptions).then((res: HttpClientResponse): Promise<any> => {
        return res.json();
      });
    const element: React.ReactElement<IHelloWorldProps> = React.createElement(
      HelloWorld,
      {
        lists: result,
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
