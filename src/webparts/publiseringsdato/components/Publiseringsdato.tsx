import * as React from 'react';
import { IPubliseringsdatoProps, ShowDates } from './IPubliseringsdatoProps';
import { SPHttpClient } from '@microsoft/sp-http';
import {
  Environment,
  EnvironmentType,
} from '@microsoft/sp-core-library';
import { Text } from 'office-ui-fabric-react/lib/Text';

export interface IPubliseringsdatoState {
  published?: Date;
  modified?: Date;
  isDraft?: boolean;
  isListItem?: boolean;
}

export interface ISPageMeta {
  Created: string;
  Modified: string;
  FirstPublishedDate?: string;
  PublishStartDate?: string;
  OData__UIVersionString: string;
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
    const {published, modified, isDraft, isListItem} = this.state;
    const {showDates, manualCreatedDate, manualModifiedDate, prefixModifiedDate} = this.props;

    const dateOptions = {year: "numeric", month: "long", day: "numeric"} as Intl.DateTimeFormatOptions;
    const locale = Intl.DateTimeFormat.supportedLocalesOf(["nb-NO", "nn-NO", "no", "da-DK", "en-US"]);

    const createdDate: Date  = manualCreatedDate && manualCreatedDate.value
      ? new Date(manualCreatedDate.value as unknown as React.ReactText)
      : published ? published : null;
    const modifiedDate: Date = manualModifiedDate && manualModifiedDate.value
      ? new Date(manualModifiedDate.value as unknown as React.ReactText)
      : modified ? modified : null;
    const showModifiedDate = showDates === ShowDates.Modified || showDates === ShowDates.Both
      || (showDates === ShowDates.Auto && (isDraft || modifiedDate && createdDate && (
        Math.abs(createdDate.getTime() - modifiedDate.getTime()) > 1000 * 60 * 5
      )));
    const showCreatedDate = showDates === ShowDates.Created || showDates === ShowDates.Both
      || (showDates === ShowDates.Auto && !isDraft && modifiedDate && createdDate && (
        createdDate > this._nDaysAgo(30) || !showModifiedDate
      ));
    return (
      <Text
        data-automation-id={`MetaDates`}
        variant={'small'}
        style={{marginTop: -12, marginBottom: -24, padding: "1px 2px 0" }}
        nowrap
        block
      >
        {showCreatedDate &&
          <span>Publisert
            {` `}
            {createdDate && <time
            data-automation-id={`CreatedDate`}
            dateTime={createdDate.toISOString()}>
            {createdDate.toLocaleDateString(locale, dateOptions)}
            {createdDate > this._nDaysAgo(1) && this._getTimeString(createdDate)}
            </time>}
          </span>
        }
        {showCreatedDate && showModifiedDate && <span>{`. `}</span> }
        {showModifiedDate &&
          <span>{isDraft ? 'Utkast oppdatert' : prefixModifiedDate}
            {` `}
            {modifiedDate && <time
              data-automation-id={`ModifiedDate`}
              dateTime={modifiedDate.toISOString()}>
              {modifiedDate.toLocaleDateString(locale, dateOptions)}
              {modifiedDate > this._nDaysAgo(1) && this._getTimeString(modifiedDate)}
            </time>}
          </span>
        }
        {showCreatedDate && showModifiedDate && <span>{`.`}</span> }
        {isListItem === false && <span>Utkast opprettet {new Date().toLocaleDateString(locale, dateOptions)}</span>}
      </Text>
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
        list: {serverRelativeUrl},
        listItem
      } = this.props.context.pageContext;
      if (!listItem) {
        this.setState({isListItem: false});
        return;
      }
      const itemId = listItem.id;
      const metaProps = ['*']; // Get all, since 'PublishStartDate' is not safe to query
      const metaPropsExpand = [];
      const url = `${absoluteUrl}/_api/web/getlist('${serverRelativeUrl}')/items(${itemId})?$select=${metaProps.join(',')}&$expand=${metaPropsExpand.join(',')}`;
      const result = await this.props.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      const meta: ISPageMeta = await result.json();
      this.setState({
        published: meta.PublishStartDate
          ? new Date(meta.PublishStartDate)
          : meta.FirstPublishedDate
            ? new Date(meta.FirstPublishedDate)
            : new Date(meta.Created),
        modified: new Date(meta.Modified),
        isDraft: !this._strEndsWith(meta.OData__UIVersionString, '.0'),
        isListItem: true,
      });
    }
  }

  private _getTimeString(date: Date): string {
    if (date.getHours() === 0 && date.getMinutes() === 0 ) return '';
    return ` kl. ${(`0${date.getHours()}`).slice(-2)}.${(`0${date.getMinutes()}`).slice(-2)}`;
  }

  private _nDaysAgo(n: number): Date {
    return new Date(new Date().getTime() - (n * 24 * 60 * 60 * 1000));
  }

  private _strEndsWith(haystack: string, needle: string): boolean {
    return haystack.substring(haystack.length - needle.length, haystack.length) === needle;
  }
}
