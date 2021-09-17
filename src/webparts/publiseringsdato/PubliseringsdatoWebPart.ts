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
import { IFile } from '@pnp/sp/files/types';

export interface IPubliseringsdatoWebPartProps {
  manualCreatedDate: IDateTimeFieldValue;
  manualModifiedDate: IDateTimeFieldValue;
  prefixModifiedDate: string;
  showDates: ShowDates;
}

export interface ISitePageDates {
  created: Date;
  modified: Date;
  firstPublished?: Date;
}

export default class PubliseringsdatoWebPart extends BaseClientSideWebPart<IPubliseringsdatoWebPartProps> {

  protected file?: IFile;
  protected dates?: ISitePageDates;
  protected isNew = true;
  protected isDraft = false;
  protected unpublishButtonPressed = false;

  public async onInit() {
    await this._updateContext();
    super.onInit();
  }

  public async onPropertyPaneConfigurationStart() {
    await this._updateContext();
    super.onPropertyPaneConfigurationStart();
  }

  private async _updateContext() {
    sp.setup(this.context);
    const pageRelativeUrl = window.location.pathname; // Can't use serverRequestPath from pageContext because of newpage bug (new context is not initialized)
    try {
      this.file = sp.web.getFileByServerRelativeUrl(pageRelativeUrl);
      const allFields = await this.file.expand('ListItemAllFields').get();
      this.isNew = allFields.MajorVersion === 0;
      this.isDraft = allFields.MinorVersion !== 0;
      this.dates = {
        created: new Date(allFields['ListItemAllFields']['Created']),
        modified: new Date(allFields['ListItemAllFields']['Modified']),
        firstPublished: allFields['ListItemAllFields']['FirstPublishedDate'] && new Date(allFields['ListItemAllFields']['FirstPublishedDate']),
      };
    } catch {}
  }

  public render(): void {
    const {manualCreatedDate, manualModifiedDate} = this.properties;
    const element: React.ReactElement<IPubliseringsdatoProps> = React.createElement(
      PubliseringsDato,
      {
        ...this.properties,
        publishedDate: manualCreatedDate && manualCreatedDate.value
          ? new Date(manualCreatedDate.value as unknown as React.ReactText)
          : this.dates ? this.dates.firstPublished || this.dates.created : undefined,
        modifiedDate: manualModifiedDate && manualModifiedDate.value
          ? new Date(manualModifiedDate.value as unknown as React.ReactText)
          : this.dates ? this.dates.modified : undefined,
        isDraft: this.isDraft,
        displayMode: this.displayMode,
      }
    );
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  private _dateToDateField(date: Date): IDateTimeFieldValue | undefined {
    if (date) return {
      value: date,
      displayValue: date.toLocaleString(),
    };
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
                groupName: `Tilpasninger${this.isNew && ' (publiser siden først!)'}`,
                groupFields: [
                  PropertyFieldDateTimePicker('manualCreatedDate', {
                    label: 'Rediger publiseringsdato og -klokkeslett',
                    disabled: this.isNew,
                    initialDate: this.properties.manualCreatedDate
                      || (this.dates && this._dateToDateField(this.dates.firstPublished))
                      || (this.dates && this._dateToDateField(this.dates.created)),
                    dateConvention: DateConvention.DateTime,
                    onPropertyChange: async (propertyPath, oldValue, newValue) => {
                      const newDate: Date = newValue.value;
                      const item = await this.file.getItem();
                      await item.validateUpdateListItem([{
                        FieldName: "FirstPublishedDate",
                        FieldValue: `${newDate.toLocaleDateString()} ${newDate.toLocaleTimeString()}`
                      }]);
                      this.onPropertyPaneFieldChanged(propertyPath, oldValue, false);
                    },
                    properties: this.properties,
                    onGetErrorMessage: null,
                    deferredValidationTime: 0,
                    key: 'manualCreatedDate',
                    showLabels: false
                  }),
                  PropertyFieldDateTimePicker('manualModifiedDate', {
                    label: 'Overstyr oppdatert-dato',
                    disabled: this.isNew,
                    initialDate: this.properties.manualModifiedDate,
                    dateConvention: DateConvention.DateTime,
                    onPropertyChange: this.onPropertyPaneFieldChanged,
                    properties: this.properties,
                    onGetErrorMessage: null,
                    deferredValidationTime: 0,
                    key: 'manualModifiedDate',
                    showLabels: false
                  }),
                  PropertyPaneButton('manualModifiedDate', {
                    text: `Bruk automatisk dato${this.dates && this.dates.modified && ` (${this.dates.modified.toLocaleDateString()})`}`,
                    disabled: this.isNew,
                    onClick: (value: any) => {
                      this.properties.manualModifiedDate = null;
                      this.context.propertyPane.close();
                      this.context.propertyPane.open();
                    },
                  }),
                  PropertyPaneChoiceGroup('prefixModifiedDate', {
                    label: 'Prefiks foran oppdatert-dato',
                    options: [
                      { key: ModifiedPrefix.Updated, text: ModifiedPrefix.Updated, disabled: this.isNew },
                      { key: ModifiedPrefix.Revised, text: ModifiedPrefix.Revised, disabled: this.isNew },
                    ],
                  }),
                ],
              },
              {
                groupName: 'Ekstra verktøy',
                groupFields: [
                  PropertyPaneButton('unpublish',{
                    text: 'Lagre og avpubliser denne siden',
                    disabled: this.isNew || this.unpublishButtonPressed,
                    onClick: async () => {
                      await this.file.checkin();
                      await this.file.unpublish('Avpublisert');
                      await this.file.checkout();
                      this.unpublishButtonPressed = true;
                      this.context.propertyPane.close();
                      this.context.propertyPane.open();
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
