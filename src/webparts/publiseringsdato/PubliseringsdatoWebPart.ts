import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  IPropertyPaneConfiguration,
  PropertyPaneButton,
  PropertyPaneChoiceGroup,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import PubliseringsDato from './components/Publiseringsdato';
import { IPubliseringsdatoProps, ShowDates, ModifiedPrefix } from './components/IPubliseringsdatoProps';
import { PropertyFieldDateTimePicker, DateConvention } from '@pnp/spfx-property-controls/lib/PropertyFieldDateTimePicker';
import { IDateTimeFieldValue } from "@pnp/spfx-property-controls/lib/PropertyFieldDateTimePicker";

export interface IPubliseringsdatoWebPartProps {
  manualCreatedDate: IDateTimeFieldValue;
  manualModifiedDate: IDateTimeFieldValue;
  prefixModifiedDate: string;
  showDates: ShowDates;
}

export default class PubliseringsdatoWebPart extends BaseClientSideWebPart<IPubliseringsdatoWebPartProps> {

  public render(): void {

    const element: React.ReactElement<IPubliseringsdatoProps> = React.createElement(
      PubliseringsDato,
      {
        context: this.context,
        ...this.properties,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'Nettdel som gjør at du kan vise sist oppdatert-dato og/eller publiseringsdato, og overstyre datoene ved behov.'
          },
          groups: [
            {
              // groupName: 'Vis hvilke datoer?',
              groupFields: [
                PropertyPaneChoiceGroup('showDates', {
                  label: 'Vis hvilke datoer?',
                  options: [
                    { key: ShowDates.Created, text: 'Publisert' },
                    { key: ShowDates.Modified, text: 'Oppdatert' },
                    { key: ShowDates.Both, text: 'Publisert og oppdatert' },
                  ],
                }),
              ],
            },
            {
              groupName: 'Tilpasninger',
              groupFields: [
                PropertyFieldDateTimePicker('manualCreatedDate', {
                  label: 'Overstyr publiseringsdato',
                  initialDate: this.properties.manualCreatedDate,
                  dateConvention: DateConvention.Date,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'manualCreatedDate',
                  showLabels: false
                }),
                PropertyPaneButton('manualCreatedDate', {
                  text: 'Fjern manuell dato',
                  onClick: (value: any) => { 
                    this.properties.manualCreatedDate = null;
                    this.context.propertyPane.close();
                    this.context.propertyPane.open();
                  },
                }),
                PropertyFieldDateTimePicker('manualModifiedDate', {
                  label: 'Overstyr oppdatert-dato',
                  initialDate: this.properties.manualModifiedDate,
                  dateConvention: DateConvention.Date,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'manualModifiedDate',
                  showLabels: false
                }),
                PropertyPaneButton('manualModifiedDate', {
                  text: 'Fjern manuell dato',
                  onClick: (value: any) => { 
                    this.properties.manualModifiedDate = null;
                    this.context.propertyPane.close();
                    this.context.propertyPane.open();
                  },
                }),
                PropertyPaneChoiceGroup('prefixModifiedDate', {
                  label: 'Prefiks foran oppdatert-dato',
                  options: [
                    { key: ModifiedPrefix.Updated, text: ModifiedPrefix.Updated },
                    { key: ModifiedPrefix.Revised, text: ModifiedPrefix.Revised },
                  ],
                }),
              ]
            },
          ]
        }
      ]
    };
  }
}
