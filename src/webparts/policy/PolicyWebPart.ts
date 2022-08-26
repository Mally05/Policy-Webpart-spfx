import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneTextField,
  PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'PolicyWebPartStrings';
import Policy from './components/Policy';
import { IPolicyProps } from './components/IPolicyProps';
import {sp} from '@pnp/sp';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import { PropertyFieldMultiSelect } from '@pnp/spfx-property-controls/lib/PropertyFieldMultiSelect';
import { SPService } from '../service/Service';

export interface IPolicyWebPartProps {
  description: string;
  lists: string;
  fields: string[];
  context: WebPartContext;
}

export default class PolicyWebPart extends BaseClientSideWebPart<IPolicyWebPartProps> {
  private _services: SPService = null;
  private _listFields: IPropertyPaneDropdownOption[] = [];


  protected onInit(): Promise<void> {
    return super.onInit().then(_ => {
      sp.setup({
        spfxContext: this.context
      });
      this.context.propertyPane.open();
      this._services = new SPService(this.context);
      this.getListFields = this.getListFields.bind(this);
    });
  }

  public render(): void {
    const element: React.ReactElement<IPolicyProps> = React.createElement(
      Policy,
      {
        context: this.context,
        description: this.properties.description,
        list: this.properties.lists,
        fields: this.properties.fields
      }
    );
    ReactDom.render(element, this.domElement);
  }



  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  public async getListFields() {
    if (this.properties.lists) {
      let allFields = await this._services.getFields(this.properties.lists);
      (this._listFields as []).length = 0;
      const x = this._listFields.push(...allFields.map(field => ({ key: field.InternalName, text: field.Title})));
      console.log(x);
    }
  }

  private listConfigurationChanged(propertyPath: string, oldValue: any, newValue: any) {
    console.log("LIST FIELDS:", this._listFields);
    if (propertyPath === 'lists' && newValue) {
      this.properties.fields = [];
      this.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
      this.context.propertyPane.refresh();
    }
    else {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    this.getListFields();
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyFieldListPicker('lists', {
                  label: 'Select a list',
                  selectedList: this.properties.lists,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  baseTemplate: 100,
                  onPropertyChange: this.listConfigurationChanged.bind(this),
                  properties: this.properties,
                  context: this.context as any,
                  onGetErrorMessage: null,
                  key: 'listPickerFieldId',
                }),PropertyPaneDropdown("color", {
                  label: "Color",
                  options: [
                    { key: "Red", text: "Red" },
                    { key: "Green", text: "Green" },
                    { key: "Blue", text: "Blue" }
                  ]
                }),
                PropertyFieldMultiSelect('fields', {
                  key: 'multiSelect',
                  label: "Multi select list fields",
                  options: this._listFields,
                  selectedKeys: this.properties.fields
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
