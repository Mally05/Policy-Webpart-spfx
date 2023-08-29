import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneButton,
  PropertyPaneButtonType,
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
import {HttpClient } from '@microsoft/sp-http';
import {
  IPropertyFieldGroupOrPerson,
  PropertyFieldPeoplePicker,
  PrincipalType
} from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';
import { IPolicyWebPartProps } from "./components/IPolicyWebPartProps";


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
  public siteName:string;
  public userGroups: any = [];

  protected onInit(): Promise<void> {
    
    return super.onInit().then((_) => {
      this.getAzureBlobWhitelistFile(this.context)
      .then((x:boolean) => {
        return this.properties.hasLicence = x;
      });

      sp.setup(this.context);
     
      this._themeProvider = this.context.serviceScope.consume(
        ThemeProvider.serviceKey
      );

      // console.log("People", this.properties.people);

      this._themeVariant = this._themeProvider.tryGetTheme();
      this._themeProvider.themeChangedEvent.add(
        this,
        this._handleThemeChangeEvent);
        this._services = new SPService();
        // this._services.getGroupMembers(this.context,this.properties.siteCollection, this.properties.people);
        // const listname = this.getListName(this.properties.listName);
        // this.loadUserGroups();
        this.getAzureBlobWhitelistFile(this.context);
        // this._services.setListPermissions(this.properties.siteCollection,listname);
        this.getLists();
      this.context.propertyPane.open();
    });
  }
  public async getAzureBlobWhitelistFile(context: WebPartContext){
    let currentTenant = context.pageContext.site.absoluteUrl;
    let domain = new URL(currentTenant);

    const result = await context.httpClient
    .get("https://signoffappgroup8c36.blob.core.windows.net/licences/Whitelist.json",
    HttpClient.configurations.v1);

    const resJson = await result.json();
    const whiteList:Array<any> = await resJson;
     
    return whiteList.some(x =>  domain.hostname == x.url ? this.properties.hasLicence = true : this.properties.hasLicence = false);
  }

  // protected async loadUserGroups(){
  //  const groups = await this._services.getListOfgroups(this.context,this.properties.siteCollection);
  //  this.userGroups = groups;

  //  console.log("OnInit: ", this.userGroups);
  // }

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
        fields: this.properties.fields,
        listName: this.properties.listName,
        isConfigured:this.configured,
        isChecked: this.isChecked,
        titleText:this.properties.titleText,
        themeVariant: this._themeVariant,
        dateSigned: this.dateSigned,
        siteName: this.properties.siteName,
        checkboxLabel: this.properties.checkboxLabel,
        hasLicence:this.properties.hasLicence,
        people:this.properties.people
      }
    );

    ReactDom.render(element, this.domElement);
    }

     private getLists = () =>{
      if(this.properties.siteCollection !== undefined){
        this._services.loadSiteCollections();
        this._services.loadDropdownListOptions(this.properties.context, this.properties.siteCollection);
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
  
  public getListName(listId:string){  
    const listName = this.properties.fields.filter(item => item.key ==listId ).map( x => {
          return x.text;
     });
     return listName[0];
  }

  protected async onPropertyPaneFieldChanged(propertyPath: string, oldValue?: any, newValue?: any) {
      if (propertyPath === 'siteCollection' && newValue) {
        this.loadingIndicator = true;
        this._services.loadDropdownListOptions(this.context,newValue).then(x =>{
          this.properties.fields = x;
          this.isFetched = false;
          super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
           this.context.propertyPane.refresh();
           this.loadingIndicator =false;
        });
      this.context.propertyPane.refresh();
    } else if 
    (propertyPath === 'listName' && newValue ) {
      const newListName = this.getListName(newValue);
      const oldListName = this.getListName(oldValue);
      this.properties.listName = newValue;
      const item = await this._services.getItemsFromSelectedList(this.context, newListName,this.properties.siteCollection);
      super.onPropertyPaneFieldChanged(propertyPath, oldListName, newListName);
      this.dateSigned = item.modified;
      this.isChecked = item.checked === undefined ? item : item.checked;
      this.configured = item.checked === undefined ? item : item.checked;
      this.context.propertyPane.refresh();
    }else if (propertyPath === 'people' && newValue){
      this.properties.people = newValue;
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
                  options: this.properties.fields,
                  disabled: this.isFetched || this.loadingIndicator,
                  selectedKey: this.properties.listName,
                }),
                // PropertyFieldPeoplePicker('people',{
                //   label: 'Add people or group to list',
                //   initialData: this.properties.people,
                //   allowDuplicate: false,
                //   principalType: [PrincipalType.Users, PrincipalType.SharePoint,PrincipalType.Security],
                //   onPropertyChange: this.onPropertyPaneFieldChanged,
                //   context: this.context,
                //   properties: this.properties,
                //   onGetErrorMessage: null,
                //   deferredValidationTime: 0,
                //   key: 'peopleFieldId'
                // }),
                PropertyPaneButton("button",{
                  onClick:null,
                  text:"Save",
                  buttonType: PropertyPaneButtonType.Primary
                })
              ],
            }
          ],
        }
      ],
    };
  }

  protected async onPropertyPaneConfigurationStart(): Promise<void> {

    if(this.properties.siteCollection !== undefined){
      const resultLists = await this._services.loadDropdownListOptions(this.context,this.properties.siteCollection);
      const response = await this._services.loadSiteCollections();
      
      this.properties.fields = resultLists;
      this.siteCollections = response;

      this.loadingIndicator =false;
      this.isFetched = false;

    } else if(this.loadingIndicator){
        const response = await this._services.loadSiteCollections();
        this.siteCollections = response;
        this.loadingIndicator = false;
      } 
      
    this.context.propertyPane.refresh();
    this.render();
  }
}
