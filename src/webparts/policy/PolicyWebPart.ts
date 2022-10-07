import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneButton,
} from "@microsoft/sp-property-pane";
import {
  BaseClientSideWebPart,
  WebPartContext,
} from "@microsoft/sp-webpart-base";
import { IODataList } from "@microsoft/sp-odata-types";
import * as strings from "PolicyWebPartStrings";
import Policy from "./components/Policy";
import { sp } from "@pnp/sp";
import  {SPService}  from "../service/Service";
import { has } from "@microsoft/sp-lodash-subset";
import { IReadonlyTheme, ThemeChangedEventArgs, ThemeProvider } from '@microsoft/sp-component-base';
import { ThemeSettingName } from "office-ui-fabric-react";
import _onconfigure from './components/Policy';

export interface IPolicyWebPartProps {
  description: string;
  lists: string;
  fields: any[];
  context: WebPartContext;
  listName: string ;
  isConfigured:boolean;
  isChecked:boolean;
  titleText:string;
  themeVariant: IReadonlyTheme | undefined;
  dateSigned: any;
}

export default class PolicyWebPart extends BaseClientSideWebPart<IPolicyWebPartProps> {
  private _themeProvider: ThemeProvider;
  private _themeVariant: IReadonlyTheme | undefined;
  private loadingIndicator:boolean = true;
  private _services: SPService;
  public _listFields: IPropertyPaneDropdownOption[] = [];
  private isFetched: boolean = false;
  protected _relativeEndUrl: string;
  public tenantName: string;
  public configured:boolean = false;
  public isChecked:boolean;
  public hasApproved:boolean;
  public displayName:string;
  public dateSigned: any;
  public listName: string;

  protected onInit(): Promise<void> {
    return super.onInit().then((_) => {

    this.listName =this.getListName(this.properties.listName)
    console.log("listname OnInit(): ", this.listName);
    
    sp.setup(this.context);
    this._themeProvider = this.context.serviceScope.consume(
      ThemeProvider.serviceKey
    );

    this._themeVariant = this._themeProvider.tryGetTheme();
    this._themeProvider.themeChangedEvent.add(
      this,
      this._handleThemeChangeEvent 
    );
    this._services = new SPService();

      this.context.propertyPane.open();
    });
  }

  private _handleThemeChangeEvent(args: ThemeChangedEventArgs): void {
    this._themeVariant = args.theme;
    this.render();
  }

  public render(): void {
    const element: React.ReactElement<IPolicyWebPartProps> = React.createElement(
      Policy,
      {
        context: this.context,
        description: this.properties.description,
        lists: this.properties.lists,
        fields: this._listFields,
        listName: this.properties.listName,
        isConfigured:this.configured,
        isChecked: this.isChecked,
        titleText:this.properties.titleText,
        themeVariant: this._themeVariant,
        dateSigned: this.dateSigned
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
   
    const listname = this.getListName(this.properties.listName);
    this._services.hasApprovedPolicyForSelectedList(this.context, listname);
    ReactDom.unmountComponentAtNode(this.domElement);
  }
  
  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  public onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    const { semanticColors
    } = currentTheme;
    this.domElement.style.setProperty('--backgroundBody', semanticColors.bodyBackground);
    this.domElement.style.setProperty('--bodyText', semanticColors.bodyText);
    this.domElement.style.setProperty('--link', semanticColors.link);
    this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered);

  }

  public getListName(listId:string){  
    const listName = this._listFields.filter(item => item.key ==listId ).map( x => {
          return x.text;
     });
     return listName[0]
  }

  protected async onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any) {
    if (propertyPath === 'listName' && newValue) {
      const newListName = this.getListName(newValue);
      const oldListName = this.getListName(oldValue);
      const item = await this._services.hasApprovedPolicyForSelectedList(this.context, newListName);
      this.dateSigned = item.modified;
      this.isChecked = item.checked === undefined ? item : item.checked;
      this.configured = item.isConfigured === undefined ? item : item.isConfigured;
      this.context.propertyPane.refresh();
      console.log(oldListName)
      super.onPropertyPaneFieldChanged(propertyPath, oldListName, newListName);

    }
    else {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, oldValue);
    }
  }

  //TODO: Loading indicator

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {

    return {
      showLoadingIndicator:this.loadingIndicator,
      loadingIndicatorDelayTime: 1,
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            { 
              groupFields: [
                PropertyPaneTextField("titleText", {
                  label: strings.Title,
                }),
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel,
                  multiline:true
                }),
                PropertyPaneDropdown("listName", {
                  label: strings.ListFieldLabel,
                  options: this._listFields,
                  selectedKey: this.properties.listName,
                  disabled:this.loadingIndicator
                }),
              ],
            },
          ],
        },
      ],
    };
  }

  
  protected onPropertyPaneConfigurationStart(): void {
    if(!this.isFetched){
      this._services.loadDropdownListOptions(this.context).then((response)=>{
        this._listFields = response;
        this.isFetched = true;
        this.loadingIndicator = false;
        this.context.propertyPane.refresh();          
      });
    }
    this.render();
  }
}
