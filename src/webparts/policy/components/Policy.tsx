import * as React from 'react';
import { IPolicyProps } from './IPolicyProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class Policy extends React.Component<IPolicyProps, {}> {
  public render(): React.ReactElement<IPolicyProps> {

    return (
      <div className="policy">
      <p>Selected List: {this.props.list}</p>
      {/* <p>Selected Fields: {this.props.fields}</p> */}
    </div>
    );
  }
}
