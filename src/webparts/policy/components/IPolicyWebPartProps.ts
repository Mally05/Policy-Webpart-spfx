import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme} from '@microsoft/sp-component-base';
import { IPropertyFieldGroupOrPerson, IPropertyFieldPeoplePickerProps } from "@pnp/spfx-property-controls";

 export interface IPolicyWebPartProps {
    siteName:string;
    siteCollection: WebPartContext;
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
    checkboxLabel:string;
    CheckboxPlaceholder:string;
    hasLicence:boolean;
    people: IPropertyFieldGroupOrPerson[];
  }