import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IDateTimeFieldValue } from "@pnp/spfx-property-controls/lib/PropertyFieldDateTimePicker";

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
  context: WebPartContext;
  manualModifiedDate: IDateTimeFieldValue;
  prefixModifiedDate: string;
  manualCreatedDate: IDateTimeFieldValue;
  showDates: ShowDates;
}
