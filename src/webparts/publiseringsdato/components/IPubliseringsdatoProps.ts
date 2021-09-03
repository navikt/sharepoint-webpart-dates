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
  modifiedDate: Date;
  isDraft?: boolean;
  version?: string;
}
