import * as React from 'react';
import styles from './HelloWorld.module.scss';
import { IHelloWorldProps } from './IHelloWorldProps';
import { escape } from '@microsoft/sp-lodash-subset';

export interface IHelloWorldState {
  items: any;
}

export default class HelloWorld extends React.Component<IHelloWorldProps, IHelloWorldState> {
  constructor(props: IHelloWorldProps) {
    super(props);
  }

  public render(): React.ReactElement<IHelloWorldProps> {
    return (
      <div className={styles.helloWorld}>
        <div className="ms-Grid">
          <div className="ms-Grid-row">
            <div className="ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12">
            <h1>My lists, Yo!</h1>
              <ul>
                {this.props.lists.map(list => 
                  <li>{list.Title}</li>
                )}
              </ul>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
