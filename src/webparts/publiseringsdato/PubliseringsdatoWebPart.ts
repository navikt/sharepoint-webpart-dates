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
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/files/web";

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
    sp.setup(this.context);
    const pageRelativeUrl = this.context.pageContext.site.serverRequestPath;
    const file = sp.web.getFileByServerRelativePath(pageRelativeUrl);

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
                    { key: ShowDates.Auto, text: 'Automatisk' },
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
                  label: 'Rediger publiseringsdato og -klokkeslett',
                  initialDate: this.properties.manualCreatedDate,
                  dateConvention: DateConvention.DateTime,
                  onPropertyChange: async (propertyPath, oldValue, newValue) => {
                    const fields = await file.expand('ListItemAllFields').get();
                    if (fields['ListItemAllFields'] && fields['ListItemAllFields']["FirstPublishedDate"]) {
                      const newDate: Date = newValue.value;
                      const item = await file.getItem();
                      await item.validateUpdateListItem([{
                        FieldName: "FirstPublishedDate",
                        FieldValue: `${newDate.toLocaleDateString()} ${newDate.toLocaleTimeString()}`
                      }]);
                    }
                    this.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
                  },
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'manualCreatedDate',
                  showLabels: false
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
              ],
            },
            {
              groupName: 'Ekstra verktøy',
              groupFields: [
                PropertyPaneButton('unpublish',{
                  text: 'Avpubliser denne siden',
                  onClick: async () => {
                    await file.checkin();
                    await file.unpublish('Avpublisert');
                  },
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
