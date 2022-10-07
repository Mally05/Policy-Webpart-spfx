import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme} from '@microsoft/sp-component-base';

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