import { WebPartContext } from "@microsoft/sp-webpart-base";
import { sp,SearchResults, Items } from "@pnp/sp/presets/all";
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
import { Web } from '@pnp/sp/presets/all';
import PolicyWebPart from "../policy/PolicyWebPart";
import { AadHttpClient, HttpClientResponse } from '@microsoft/sp-http';

  export class SPService {
  
  private web = Web("https://connectkonicaminolta.sharepoint.com/sites/BSW_BPS");  
   


    public async getSiteCollections() {
      try {
       const item = await sp.search({
          Querytext: "contentclass:STS_Site",
          SelectProperties: ["Title", "SPSiteUrl"],
          RowLimit: 500,
          TrimDuplicates: false,
        })
        .then((searchResults: SearchResults) => {
          return searchResults.PrimarySearchResults
        });
        return item

      } catch (error) {
        console.log("Couldn't fetch Site collection data: ", error)
      }
    }

    public async getLists(context: WebPartContext,url?) {
      const uri = `/_api/web/lists?$filter=BaseType ne 1 and Hidden eq false`;
      try {
        return await context.spHttpClient
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

    public async getListsForSelectedSites(context: WebPartContext,url?:WebPartContext) {
      const uri = `${url}/_api/web/lists?$filter=BaseType ne 1 and Hidden eq false`;
      try {
        return await context.spHttpClient
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

    public async getItemsFromSelectedList(context: WebPartContext,listName: string, url?:any) {
      const uri = `${url}/_api/web/lists/GetByTitle('${listName}')/items?$select=Id,Approved,Modified,Employee/EMail&$expand=Employee/Id`;
      try {
        return await context.spHttpClient
          .get(uri, SPHttpClient.configurations.v1)
          .then((res: SPHttpClientResponse) => {
            if (res.ok) {
              const item = {
                checked: false,
                isConfigured:true,
                res: res.json()
              }
              return item;

            } else {
              const item: any = {
                checked: false,
                isConfigured: true,
              };
              console.log(`There was an error: \n Can't find columns: Approved, Modified, Employee in ${listName}`);
              return item
            }
          });
      } catch (error) {
        console.log(`Error: ${error}`);
      }
      return Error("Bad Request");
    }

    public hasApprovedSelectedList = async (context: WebPartContext,listName: string, url?:any) => {
      
      const hasApproved  = await this.getItemsFromSelectedList(context,listName,url).then((x: any) => {
       const result = x.res.then((item:any)=>{
          const currentUserEmail = context.pageContext.user.email;
          const convertToArray = [];
          convertToArray.push(item);
          const spItem = convertToArray[0].value.filter((x: any) => currentUserEmail == x.Employee.EMail);
  
          if (spItem.length > 0) {
            for (var i = 0; i < spItem.length; i++) {
              let items = spItem[i];
              if (items.Employee.EMail == currentUserEmail && items.Approved == true) {
                console.log("You have accepted this policy");
                const date = new Date(items.Modified);
                const spObj = {
                  checked: true,
                  modified: date.toDateString(),
                };
                return spObj;

              } else if (items.Employee.EMail == currentUserEmail && items.Approved == false) {
                console.log("You have not accepted this policy");
                const item: any = {
                  checked: false,
                  isConfigured: true,
                };
                return item;

              } else if (items.Employee.EMail !== "") {
                const item: any = {
                  checked: false,
                  isConfigured: true,
                };
                 console.error("We couldn't find you in the list please make sure you are added to the selected policy list.");

                return item
              }
            }
          }else {
            const item: any = {checked: false,isConfigured: true};
             console.error("We couldn't find you in the list please make sure you are added to the selected policy list.");
            return item
          }
        })
        return result;
      }).catch( () =>{
        const item: any = {
          checked: false,
          isConfigured: true,
        };
        return item;
      });
      console.log("Approved??",hasApproved);
      return hasApproved;
    };
    public async loadSiteCollections() {
     return await this.getSiteCollections().then((results: any) => {
          const arr: Array<IPropertyPaneDropdownOption> =
          new Array<IPropertyPaneDropdownOption>();
          results.map((col: any) => {
            arr.push({
              key:  col.SPSiteUrl,  
              text: col.Title
            });
          });
          return arr;
        });
    }
    public async loadDropdownListOptions(context:WebPartContext,url?) {
      return await this.getListsForSelectedSites(context,url).then((response: any) => {
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

    public async patchItemToSharePoint({listName,Status,context,_spHttpContext,currentUser,siteCollection}) {
      const spitems = await this.getItemsFromSelectedList(context, listName,siteCollection);
      const result = await spitems.res.catch((error:any) =>{ console.error(error)});
      const filteredResult = await result.value.filter((spItem: any) => currentUser == spItem.Employee.EMail);   

      if(filteredResult.length <= 0){
        return console.error("Failed to post");
      }
      
      const { Id , Modified }: any = filteredResult[0];
      const url = `${siteCollection}/_api/web/lists/getbytitle('${listName}')/items(${Id})`;
      
      const body: any = { Approved: Status};
      const spHttpClientOptions: ISPHttpClientOptions = { body: JSON.stringify(body)};

      try {
        const spPost = await _spHttpContext.post(
          url,
          SPHttpClient.configurations.v1,
          {
            headers: {
              Accept: "application/json;odata=nometadata",
              "Content-type": "application/json;odata=nometadata",
              "X-HTTP-Method": "MERGE",
              "odata-version": "",
              "IF-MATCH": "*",
            },
            body: spHttpClientOptions.body,
          }
        );

        const result: any = {
          modified: Modified,
          response: spPost,
          checked: true,
        };

        return result;
      } catch (error) {
        console.log(`Error: ${error}`);
      }
    }

    // public async patchTenantIdTolist(context: WebPartContext){
    //   const listname = "Sign Off Whitelist";
    // try {
    //      await this.web.lists.getByTitle(listname).items.add({
    //       SiteURL: context.pageContext.site.absoluteUrl
    //     })
    //     console.log("first patch succeded!")
        
    //   } catch (error) {
    //     console.log("This keeps failing for some reason :(",error)
    //   }
    // }
  }
