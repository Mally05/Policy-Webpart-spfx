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
import * as strings from "PolicyWebPartStrings";
import Policy from "./components/Policy";
import { sp } from "@pnp/sp";
import  {SPService}  from "../service/Service";
import { IReadonlyTheme, ThemeChangedEventArgs, ThemeProvider } from '@microsoft/sp-component-base';
import _onconfigure from './components/Policy';
import { AadHttpClient, HttpClientResponse,ISPHttpClientOptions } from '@microsoft/sp-http';

export interface IPolicyWebPartProps {
  siteName:string;
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
  siteCollection: any[];
  checkboxLabel: string;
}

export default class PolicyWebPart extends BaseClientSideWebPart<IPolicyWebPartProps> {
  private _themeProvider: ThemeProvider;
  private _themeVariant: IReadonlyTheme | undefined;
  private loadingIndicator:boolean = true;
  private _services: SPService;
  public _listFields: IPropertyPaneDropdownOption[] = [];
  public siteCollections: IPropertyPaneDropdownOption[] = [];
  private isFetched: boolean = true;
  protected _relativeEndUrl: string;
  public tenantName: string;
  public configured:boolean = false;
  public isChecked:boolean;
  public hasApproved:boolean;
  public displayName:string;
  public dateSigned: any;
  public listName: string;
  public siteName:string

  protected onInit(): Promise<void> {
    return super.onInit().then((_) => {

      console.log("ClientID: ",this.context.pageContext)
      console.log("Site: ",this.context.pageContext.site.absoluteUrl)

      let client = new AadHttpClient(this.context.serviceScope,"639e1eed-fd88-47d1-a893-82517b799865");
      console.log("ServiceScope: ",this.context.serviceScope)
      const body:any ={
        name: this.context.pageContext.site.absoluteUrl
      }
      const aadClientOptions: ISPHttpClientOptions = {body: JSON.stringify(body)};

      client.post("https://sign-off-app.azurewebsites.net/api/demo?",AadHttpClient.configurations.v1,{body: aadClientOptions.body})
      .then((x:HttpClientResponse) =>{
        x.json().then(result => {
          console.log(result.status);
        }).catch(err => {
          console.error(err)
        })
      }).catch(error =>{
        console.log(error);
      })

    sp.setup(this.context);
    this._themeProvider = this.context.serviceScope.consume(
      ThemeProvider.serviceKey
      );
      
      this._themeVariant = this._themeProvider.tryGetTheme();
      this._themeProvider.themeChangedEvent.add(
        this,
        this._handleThemeChangeEvent);
        this._services = new SPService();
        this.getLists();
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
        siteCollection:this.properties.siteCollection,
        context: this.context,
        description: this.properties.description,
        lists: this.properties.lists,
        fields: this._listFields,
        listName: this.properties.listName,
        isConfigured:this.configured,
        isChecked: this.isChecked,
        titleText:this.properties.titleText,
        themeVariant: this._themeVariant,
        dateSigned: this.dateSigned,
        siteName: this.properties.siteName,
        checkboxLabel: this.properties.checkboxLabel
      }
    );

    ReactDom.render(element, this.domElement);
  }

     private getLists = () =>{
      if(this.properties.siteCollection !== undefined){
        this._services.loadDropdownListOptions(this.properties.context, this.properties.siteCollection);
        this._services.loadSiteCollections();
      }else {
        this._services.loadSiteCollections();
      }
    }

    protected onDispose(): void {
   
    const listname = this.getListName(this.properties.listName);
    this._services.getItemsFromSelectedList(this.context, listname);
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
  public async loadDropdown(){
      
  }

  public getListName(listId:string){  
    const listName = this._listFields.filter(item => item.key ==listId ).map( x => {
          return x.text;
     });
     return listName[0]
  }

  protected async onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any) {
      if (propertyPath === 'siteCollection' && newValue) {
        this.loadingIndicator = true;
        this._services.loadDropdownListOptions(this.context,newValue).then(x =>{
          this._listFields = x;
          this.isFetched = false;
          super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
           this.context.propertyPane.refresh();
           this.loadingIndicator =false;
        })
      this.context.propertyPane.refresh();
    } else if 
    (propertyPath === 'listName' && newValue ) {
      const newListName = this.getListName(newValue);
      const oldListName = this.getListName(oldValue);
      const item = await this._services.getItemsFromSelectedList(this.context, newListName,this.properties.siteCollection);
      super.onPropertyPaneFieldChanged(propertyPath, oldListName, newListName);
      this.dateSigned = item.modified;
      this.isChecked = item.checked === undefined ? item : item.checked;
      this.configured = item.checked === undefined ? item : item.checked;
      this.context.propertyPane.refresh();
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
                  placeholder: strings.TitlePlaceholder
                }),
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel,
                  multiline:true,
                  placeholder:strings.Description
                }),
                PropertyPaneTextField("checkboxLabel", {
                  label: strings.CheckboxLabel,
                  placeholder: strings.CheckboxPlaceholder
                }),
                 PropertyPaneDropdown("siteCollection", {
                  label: strings.SiteCollection,
                  options: this.siteCollections,
                  selectedKey: this.properties.siteName,
                  disabled: this.loadingIndicator
                }),
                PropertyPaneDropdown("listName",{
                  label: strings.ListFieldLabel,
                  options: this._listFields,
                  disabled: this.isFetched || this.loadingIndicator,
                  selectedKey: this.properties.listName,
                })
              ],
            }
          ],
        },
      ],
    };
  }

  protected async onPropertyPaneConfigurationStart(): Promise<void> {
    if(this.loadingIndicator){
      const response = await this._services.loadSiteCollections();
      this.siteCollections = response;
      this.loadingIndicator = false;
    }
      this.context.propertyPane.refresh(); 
    
    this.render();
  }
}
