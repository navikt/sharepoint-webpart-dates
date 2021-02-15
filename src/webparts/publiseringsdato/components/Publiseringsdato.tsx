import * as React from 'react';
import { IPubliseringsdatoProps, ShowDates } from './IPubliseringsdatoProps';
import { SPHttpClient } from '@microsoft/sp-http';
import {
  Environment,
  EnvironmentType,
} from '@microsoft/sp-core-library';

export interface IPubliseringsdatoState {
  published?: Date;
  modified?: Date;
}

export interface ISPageMeta {
  Created: string;
  Modified: string;
  FirstPublishedDate?: string;
  PublishStartDate?: string;
}

export default class PubliseringsDato extends React.Component<IPubliseringsdatoProps, IPubliseringsdatoState> {

  public constructor(props: IPubliseringsdatoProps) {
    super(props);
    this.state = {};
  }

  public async componentDidMount() {
    await this._getPageMeta();
  }

  public render(): React.ReactElement<IPubliseringsdatoProps> {
    const {published: created, modified} = this.state;
    const {showDates, manualCreatedDate, manualModifiedDate, prefixModifiedDate} = this.props;
    const createdDate: Date  = manualCreatedDate && manualCreatedDate.value 
      ? new Date(manualCreatedDate.value as unknown as React.ReactText)
      : created ? created : null;
    const modifiedDate: Date = manualModifiedDate && manualModifiedDate.value
      ? new Date(manualModifiedDate.value as unknown as React.ReactText)
      : modified ? modified : null;
    const dateOptions = {year: 'numeric', month: 'long', day: 'numeric'};
    const showCreatedDate = showDates === ShowDates.Created || showDates === ShowDates.Both;
    const showModifiedDate = showDates === ShowDates.Modified || showDates === ShowDates.Both;
    return (
      <div data-automation-id={`MetaDates`} style={{marginTop:-12, marginBottom: -24, padding: "1px 2px 0"}}>
        {showCreatedDate &&
          <span>Publisert
            {` `}
            {createdDate && <time
            data-automation-id={`CreatedDate`}
            dateTime={createdDate.toISOString()}>
            {createdDate.toLocaleDateString(undefined, dateOptions)}
            </time>}
          </span>
        }
        {showCreatedDate && showModifiedDate && <span>{` // `}</span> }
        {showModifiedDate &&
          <span>{prefixModifiedDate}
            {` `}
            {modifiedDate && <time
              data-automation-id={`ModifiedDate`}
              dateTime={modifiedDate.toISOString()}>
              {modifiedDate.toLocaleDateString(undefined, dateOptions)}
            </time>}
          </span>
        }
      </div>
    );
  }

  private async _getPageMeta() {
    if (Environment.type === EnvironmentType.Local) {
      this.setState({
        published: new Date('2018-01-01T12:00:00Z'),
        modified: new Date(),
      });
    } else if (Environment.type === EnvironmentType.SharePoint || Environment.type === EnvironmentType.ClassicSharePoint) {
      const {
        web: {absoluteUrl},
        list: {id: listId},
        listItem: {id: ItemId}
      } = this.props.context.pageContext;
      const metaProps = []; // Get all, since 'PublishStartDate' is not safe to query
      const metaPropsExpand = [];
      const url = `${absoluteUrl}/_api/web/lists(guid'${listId}')/items(${ItemId})?$select=${metaProps.join(',')}&$expand=${metaPropsExpand.join(',')}`;
      const result = await this.props.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      const meta = await result.json();
      this.setState({
        published: meta.PublishStartDate 
          ? new Date(meta.PublishStartDate)
          : meta.FirstPublishedDate 
            ? new Date(meta.FirstPublishedDate)
            : new Date(meta.Created),
        modified: new Date(meta.Modified),
      });
    }
  }
}
