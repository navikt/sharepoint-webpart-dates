import * as React from 'react';
import { IPubliseringsdatoProps, ShowDates } from './IPubliseringsdatoProps';
import { Text } from 'office-ui-fabric-react/lib/Text';
import { DisplayMode } from '@microsoft/sp-core-library';

export default class PubliseringsDato extends React.Component<IPubliseringsdatoProps> {

  public constructor(props: IPubliseringsdatoProps) {
    super(props);
    this.state = {};
  }

  public renderDate(prefix: string, date: Date, automationId: string): React.ReactElement {
    const dateOptions = {year: "numeric", month: "long", day: "numeric"} as Intl.DateTimeFormatOptions;
    const locale = Intl.DateTimeFormat.supportedLocalesOf(["nb-NO", "nn-NO", "no", "da-DK", "en-US"]);

    return (
      <>
        {`${prefix} `}
        {date && <time
          data-automation-id={automationId}
          dateTime={date.toISOString()}>
          {date.toLocaleDateString(locale, dateOptions)}
          {this._dateIsLessThanDaysAgo(date, 1) && this._getTimeString(date)}
        </time>}
      </>
    );
  }

  public render(): React.ReactElement<IPubliseringsdatoProps> {
    const {
      showDates,
      publishedDate,
      modifiedDate,
      prefixModifiedDate,
      isDraft,
      displayMode,
    } = this.props;

    /***
     * Show created date if
     * - web part property is set to only show created date or both modified and created
     * - OR if property is set to Auto AND the page is not a draft and the publish date is less than 30 days ago.
     */
    const showCreatedDate = publishedDate && (showDates === ShowDates.Created || showDates === ShowDates.Both
      || showDates === ShowDates.Auto && publishedDate && this._dateIsLessThanDaysAgo(publishedDate, 30));
    /***
     * Show modified date if
     * - web part property is set to only show modified date or both modified and published
     * - OR if property is set to Auto AND the page is a draft OR the page was only modified less than 10 minutes after it was published
     */
     const showModifiedDate = modifiedDate && (showDates === ShowDates.Modified || showDates === ShowDates.Both || showDates === ShowDates.Auto && (
      !showCreatedDate ||
      publishedDate && (
        this._minutesBetweenDates(publishedDate, modifiedDate) > 10 && (
          !this._datesAreOnSameDay(publishedDate, modifiedDate) ||
          this._dateIsLessThanDaysAgo(publishedDate, 2)
        )
      )
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
        {showCreatedDate && showModifiedDate && `. ` }
        {showModifiedDate && this.renderDate(prefixModifiedDate, modifiedDate, 'ModifiedDate')}
        {showCreatedDate && showModifiedDate && `. ` }
        {!showCreatedDate && !showModifiedDate && displayMode == DisplayMode.Edit && 'Last inn siden på nytt for å vise publiseringsdato.'}
        {isDraft && ` (UTKAST)`}
      </Text>
    );
  }

  private _getTimeString(date: Date): string {
    if (date.getHours() === 0 && date.getMinutes() === 0 ) return '';
    return ` kl. ${(`0${date.getHours()}`).slice(-2)}.${(`0${date.getMinutes()}`).slice(-2)}`;
  }

  private _dateIsLessThanDaysAgo(date: Date, days: number): boolean {
    return date > new Date(new Date().getTime() - (days * 24 * 60 * 60 * 1000));
  }

  private _datesAreOnSameDay(first: Date, second: Date): boolean {
    return first.getFullYear() === second.getFullYear() &&
    first.getMonth() === second.getMonth() &&
    first.getDate() === second.getDate() ;
  }

  private _minutesBetweenDates(first: Date, second: Date): number {
    return Math.abs(first.getTime() - second.getTime()) / (1000 * 60);
  }

}
