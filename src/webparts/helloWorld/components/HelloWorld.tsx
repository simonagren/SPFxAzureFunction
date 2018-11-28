import * as React from 'react';
import styles from './HelloWorld.module.scss';
import { IHelloWorldProps } from './IHelloWorldProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { IHttpClientOptions, AadHttpClient, HttpClientResponse } from '@microsoft/sp-http';
import { Spinner, PrimaryButton, Label, TextField, autobind } from 'office-ui-fabric-react';

export interface IHelloWorldState {
  teamName: string;
  lists: any;
  isLoading: boolean;
}

export default class HelloWorld extends React.Component<IHelloWorldProps, IHelloWorldState> {
  constructor(props: IHelloWorldProps) {
    super(props);

    this.state = {
      lists: null,
      teamName: "",
      isLoading: false,
    };

  }

  public render(): React.ReactElement<IHelloWorldProps> {
    return (
      <div className={styles.helloWorld}>
        <div className="ms-Grid">
          <div className="ms-Grid-row">
            <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12">
              <Label>Get all lists for Site</Label>
            </div>
          </div>
          <div className="ms-Grid-row">
            <div className="ms-Grid-col ms-u-sm6 ms-u-md6 ms-u-lg6">
              <TextField
                placeholder={"Enter Site Name"}
                resizable={false}
                onChanged={(text) => this.setState({ teamName: text })}
              />
            </div>
            <div className="ms-Grid-col ms-u-sm6 ms-u-md6 ms-u-lg6">
              <PrimaryButton
                data-automation-id="test"
                disabled={!(this.state.teamName.length > 0)}
                text="Send"
                onClick={this._createTeam}
              />
            </div>
          </div>
          <div className="ms-Grid-row">
            <div className="ms-Grid-col ms-u-sm6 ms-u-md6 ms-u-lg6">
              {!this.state.lists && this.state.isLoading &&
                <Spinner></Spinner>
              }
              <ul>
              {this.state.lists && this.state.lists.map(list => 
              <li>
                {list.Title}
              </li>)
              }
              </ul>
            </div>
          </div>
        </div >
      </div >
    );
  }

  @autobind
  private async _createTeam(): Promise<void> {
    this.setState({
      isLoading: true
    });
    // Setup the options with header and body
    const headers: Headers = new Headers();
    headers.append("Content-type", "application/json");

    const postOptions: IHttpClientOptions = {
      headers: headers,
      body: `{"site": "${this.state.teamName}"}`
    };

    const result: any = await this.props.client
      .post('https://pnptesting.azurewebsites.net/api/PnPNewInvite', AadHttpClient.configurations.v1, postOptions).then((res: HttpClientResponse): Promise<any> => {
        this.setState({
          isLoading: false
        });
        return res.json();

      }).catch(error => {
        this.setState({
          isLoading: false,
        });
        console.error(error);
      }
      );
    this.setState({
      lists: result
    });
  }
}
