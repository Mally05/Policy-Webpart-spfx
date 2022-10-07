import { WebPartContext } from "@microsoft/sp-webpart-base";
import { sp } from "@pnp/sp/presets/all";
import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions,
} from "@microsoft/sp-http";
import { IPropertyPaneDropdownOption } from "@microsoft/sp-property-pane";
import { IODataList } from "@microsoft/sp-odata-types";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import PolicyWebPart from "../policy/PolicyWebPart";
import * as React from "react";
import { SpinButton } from "office-ui-fabric-react";
import { xor } from "lodash";


  export class SPService {
 
  public async getLists(context:WebPartContext) {
    const uri = `/_api/web/lists?$filter=BaseType ne 1 and Hidden eq false`;
    try {
      return context.spHttpClient
        .get(uri, SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse) => {
          if (response.ok) {
            return response.json();
          } else {
            return console.log(`There was an error fetching the URL.`);
          }
        });
    } catch (error) {
      return console.log(`Error: ${error}`);
      }
    }

  public async getItemsFromSelectedList(context: WebPartContext, listName: string) {

    const uri = `/_api/web/lists/GetByTitle('${listName}')/items?$select=Id,Policyapproved,Modified,Employee/EMail&$expand=Employee/Id`
    try {
      return await context.spHttpClient
        .get(uri, SPHttpClient.configurations.v1)
        .then((res: SPHttpClientResponse) => {
          if (res.ok) {
            return res.json();
          } else {
            console.log("There was an error");
          }
        });
    } catch (error) {
      console.log(`Error: ${error}`);
    }

    return Error("Bad Request");
  }

  public hasApprovedPolicyForSelectedList = async (context: WebPartContext, listName:string) => {
    
     const hasApproved = await this.getItemsFromSelectedList(context,listName).then((item) =>{

      const currentUserEmail = context.pageContext.user.email;
      const convertToArray = [];
      convertToArray.push(item);
      const spItem = convertToArray[0].value.filter((x:any) => currentUserEmail == x.Employee.EMail);

     if(spItem.length > 0){
       for(var i=0; i<spItem.length; i++){
        let items = spItem[i];
        if(items.Employee.EMail == currentUserEmail && items.Policyapproved == true){
          console.log("You have accepted this policy");
          const date = new Date(items.Modified);
          
          const spObj = {
            checked: true,
            modified: date.toDateString(),
          }

          return spObj;  
          } else if(items.Employee.EMail == currentUserEmail && items.Policyapproved == false){
            console.log("You have not accepted this policy");
            const item:any = {
              checked: false,
              isConfigured: true
            }
            return  item;
           }else if (items.Employee.EMail !== currentUserEmail){
            console.log("We couldn't find you in the list please make sure you are added to the selected policy list.");
          }
        }
      }
    });
    return hasApproved;
  }

  public async loadDropdownListOptions(context: WebPartContext) {
  
    return  await this.getLists(context).then((response:any) => {
      const options: Array<IPropertyPaneDropdownOption> =
        new Array<IPropertyPaneDropdownOption>();
      response.value.map((list: IODataList) => {
        options.push({
          key: list.Id,
          text: list.Title,
        });
        context.propertyPane.refresh();
      });
      return options;
    });
  }

  public async patchItemToSharePoint({listName, Status, context,_spHttpContext,currentUser}) {

    const items = await this.getItemsFromSelectedList(context,listName);
    const [{Id}] = items.value.filter( (spItem:any) => currentUser == spItem.Employee.EMail);
    const [{Modified}] = await items.value.filter( (spItem:any) => currentUser == spItem.Employee.EMail);
    
  const url = context.pageContext.site.absoluteUrl + `/_api/web/lists/getbytitle('${listName}')/items(${Id})`;
  const body: any = {
      Policyapproved: Status
    };    

    const spHttpClientOptions: ISPHttpClientOptions = { 
      body: JSON.stringify(body),
    };

    try {
        const spPost = await _spHttpContext
        .post(url, SPHttpClient.configurations.v1,{
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'Content-type': 'application/json;odata=nometadata',
              'X-HTTP-Method': 'MERGE',
              'odata-version': '',
              'IF-MATCH': '*'
            },
            body: spHttpClientOptions.body
          });

          const result:any = {
            modified:Modified,
            response: spPost,
            checked: true
          };

         return result;
    } catch (error) {
        console.log(`Error: ${error}`);
    }
  }
}
