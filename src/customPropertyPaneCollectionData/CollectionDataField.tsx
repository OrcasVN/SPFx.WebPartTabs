import * as React from 'react';
import { FieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-controls-react/lib/FieldCollectionData';

export interface ICollectionDataFieldProps {
  label: string;
  value: any[];
  onChanged: (value: any[]) => void;
  stateKey: string;
}

export interface IWebPartOption {
  key: string;
  text: string;
}

export interface ICollectionDataFieldState {
  options: IWebPartOption[];
  error: string;
}

export default class CollectionDataField extends React.Component<ICollectionDataFieldProps, ICollectionDataFieldState> {
  constructor(props: ICollectionDataFieldProps, state: ICollectionDataFieldState) {
    super(props);
    this.state = {
      options: undefined,
      error: undefined
    };
  }

  public componentDidMount(): void {
    this.loadOptions();
  }

  public componentDidUpdate(prevProps: ICollectionDataFieldProps, prevState: ICollectionDataFieldState): void {
    if (this.props.stateKey !== prevProps.stateKey) {
      this.loadOptions();
    }
  }

  private loadOptions(): void {
    this.setState({
      error: undefined,
      options: undefined
    });
    const allWebPartOnPage = document.querySelectorAll('div[data-sp-feature-instance-id*="-"]')
    let options = []
    allWebPartOnPage.forEach(item => {
      if (item.getAttribute('data-sp-feature-tag').indexOf('Web Part Tabs') < 0) {
        options.push({
          key: item.getAttribute('data-sp-feature-instance-id'),
          text: item.getAttribute('data-sp-feature-tag')
        })
      }
    })

    if (!options || options.length === 0) {
      this.setState({
        error: 'There are no webparts on this page',
        options: []
      });
      return
    }
    this.setState({
      error: undefined,
      options: options
    });
  }

  public render(): JSX.Element {
    const error: JSX.Element = this.state.error !== undefined ? <div className={'ms-TextField-errorMessage ms-u-slideDownIn20'}>Error while loading items: {this.state.error}</div> : <div />;

    return (
      <div>
        <FieldCollectionData
          key={"collectionTabs"}
          label={this.props.label}
          manageBtnLabel={"Manage Collection Tabs"}
          onChanged={this.onChanged.bind(this)}
          panelHeader={"Collection tabs"}

          fields={[
            { id: "Title", title: "Title", type: CustomCollectionFieldType.string, required: true },
            { id: "WebPart", title: 'Web Part', type: CustomCollectionFieldType.dropdown, options: this.state.options, required: true },
            { id: "DisplayOrder", title: "Display Order", type: CustomCollectionFieldType.number },
            { id: "IconName", title: "Icon Name", type: CustomCollectionFieldType.string },
          ]}
          value={this.props.value}
        />
        {error}
      </div>
    );
  }

  private onChanged(value: any[]): void {
    if (this.props.onChanged) {
      this.props.onChanged(value);
    }
  }
}