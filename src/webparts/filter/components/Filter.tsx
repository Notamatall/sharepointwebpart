import * as React from 'react';
import styles from './Filter.module.scss';
import { IFilterProps } from './IFilterProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient } from '@microsoft/sp-http';
import DetailsListDocumentsExample from './DetailsListDocumentExample';


export default class Filter extends React.Component<IFilterProps, { list: any }> {


  constructor(props: IFilterProps) {
    super(props);

    this.state = {
      list: props.list
    };
  }
  // private onGetListItemsClicked = async (event: React.MouseEvent<HTMLButtonElement>): Promise<void> => {
  //   event.preventDefault();

  //   this.setState({list:}) = await this.getListItems();
  //   console.log(this.list);
  // }


  public render(): React.ReactElement<IFilterProps> {

    // const response = this.props.context.spHttpClient.get(
    //   `https://elfodev.sharepoint.com/_api/web/lists/Items`,
    //   SPHttpClient.configurations.v1)
    //   .then(value => value.json())
    //   .then(json => { this.setState({ list: json }) }
    //   );


    return (
      <div>
        {
          this.state.list &&
          <DetailsListDocumentsExample
            list={this.state.list}
            context={this.props.context} />
        }
      </div>
    );
  }
}