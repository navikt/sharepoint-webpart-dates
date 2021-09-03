import * as React from 'react';
import { IPubliseringsdatoProps, ShowDates } from './IPubliseringsdatoProps';
import { Text } from 'office-ui-fabric-react/lib/Text';

export default class PubliseringsDato extends React.Component<IPubliseringsdatoProps> {

  public constructor(props: IPubliseringsdatoProps) {
    super(props);
    this.state = {};
  }

  public renderDate(prefix: string, date: Date, automationId: string) {
    const dateOptions = {year: "numeric", month: "long", day: "numeric"} as Intl.DateTimeFormatOptions;
    const locale = Intl.DateTimeFormat.supportedLocalesOf(["nb-NO", "nn-NO", "no", "da-DK", "en-US"]);

    return (
      <span>{prefix}
        {` `}
        {date && <time
          data-automation-id={automationId}
          dateTime={date.toISOString()}>
          {date.toLocaleDateString(locale, dateOptions)}
          {date > this._nDaysAgo(1) && this._getTimeString(date)}
        </time>}
      </span>
    );
  }

  public render(): React.ReactElement<IPubliseringsdatoProps> {
    const {
      showDates,
      publishedDate,
      modifiedDate,
      prefixModifiedDate,
      isDraft,
      version,
    } = this.props;

    const showModifiedDate = showDates === ShowDates.Modified || showDates === ShowDates.Both
      || (showDates === ShowDates.Auto && (isDraft || publishedDate && (
        Math.abs(publishedDate.getTime() - modifiedDate.getTime()) > 1000 * 60 * 5
      )));
    const showCreatedDate = publishedDate && showDates === ShowDates.Created || showDates === ShowDates.Both
      || (showDates === ShowDates.Auto && !isDraft && publishedDate && (
        publishedDate > this._nDaysAgo(30)
      ));
    return (
      <Text
        data-automation-id={`MetaDates`}
        variant={'small'}
        style={{marginTop: -12, marginBottom: -24, padding: "1px 2px 0" }}
        nowrap
        block
      >
        {showCreatedDate && this.renderDate('Publisert', publishedDate, 'CreatedDate')}
        {showCreatedDate && showModifiedDate && <span>{`. `}</span> }
        {showModifiedDate && this.renderDate(prefixModifiedDate, modifiedDate, 'ModifiedDate')}
        {showCreatedDate && showModifiedDate && <span>{`. `}</span> }
        {isDraft && ` (UTKAST${version ? `, v${version}` : ''})`}
      </Text>
    );
  }

  private _getTimeString(date: Date): string {
    if (date.getHours() === 0 && date.getMinutes() === 0 ) return '';
    return ` kl. ${(`0${date.getHours()}`).slice(-2)}.${(`0${date.getMinutes()}`).slice(-2)}`;
  }

  public _nDaysAgo(n: number): Date {
    return new Date(new Date().getTime() - (n * 24 * 60 * 60 * 1000));
  }
}
