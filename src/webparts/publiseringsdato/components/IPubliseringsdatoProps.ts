import { DisplayMode } from '@microsoft/sp-core-library';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

export enum ShowDates {
  Auto = 'Auto',
  Created = 'Created',
  Modified = 'Modified',
  Both = 'Both',
}

export enum ModifiedPrefix {
  Updated = 'Oppdatert',
  Revised = 'Gjennomgått',
}

export interface IPubliseringsdatoProps {
  showDates: ShowDates;
  prefixModifiedDate: string;
  publishedDate?: Date;
  modifiedDate?: Date;
  isDraft?: boolean;
  displayMode: DisplayMode;
  themeVariant: IReadonlyTheme;
}
