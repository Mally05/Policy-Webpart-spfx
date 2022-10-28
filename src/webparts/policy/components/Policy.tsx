import * as React from "react";
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import {
  Checkbox,
  Text,
  IStackTokens,
  ITheme,
  Stack,
  MessageBar,
  MessageBarButton,
  MessageBarType,
} from "office-ui-fabric-react";
import  {SPService}  from "../../service/Service";
import { IPolicyWebPartProps } from "./IPolicyWebPartProps";
import styles from './Policy.module.scss';
import { SPHttpClientResponse} from "@microsoft/sp-http";
import * as strings from "PolicyWebPartStrings";

interface IListProps{
  signed: any;
  listName:any;
}

interface CheckedState {
checked: boolean;
date:any
}

const SuccessExample = (p:IListProps) => (
  <MessageBar
    messageBarType={MessageBarType.success}
    isMultiline={false}
  >
    {strings.Accepted} {p.listName} {strings.On} <strong>{p.signed}</strong>
  </MessageBar>
);

 export default class Policy extends React.Component<IPolicyWebPartProps, CheckedState> {
  
   private _services: SPService;
   private isChecked: boolean;
   private listName: string;
   private dateSigned: any;

   constructor(props) {
    super(props);
    this.state = { checked: this.props.isChecked,
      date:""
    }
  }
  
   
   
componentDidUpdate(prevProps: Readonly<IPolicyWebPartProps>, prevState: Readonly<CheckedState>, snapshot?: any): void {
  if(prevProps.isChecked !== this.props.isChecked && this.props.listName !== prevProps.listName){
    const listName = this.getListName(this.props.listName);
    const result = this._services.hasApprovedSelectedList(this.props.context, listName,this.props.siteCollection)
    result.then((res:any)=>{
        this.setState({checked: res.checked, date: res.modified})
      })
  }
}

async componentDidMount(): Promise<void> {
  
  //  this._services.patchTenantIdTolist(this.props.context);

   const currentUserEmail = this.props.context.pageContext.user.email;

   const selectedSite = await this._services.getListsForSelectedSites(this.props.context, this.props.siteCollection)

   const selectedList: any = await selectedSite.value.filter((x:any) => x.Id ==this.props.listName).map((listname: any) => {return listname.Title});
   
   const itemObj = await this._services.getItemsFromSelectedList(this.props.context, selectedList[0], this.props.siteCollection);

   const hasApproved = await this._services.hasApprovedSelectedList(this.props.context,selectedList[0], this.props.siteCollection);
   
   
   const {checked, modified} = hasApproved

    if(checked){
      const date = new Date(modified);
      this.setState({checked: checked, date: date.toDateString()})
    }

   }


   public  render(): React.ReactElement<IPolicyWebPartProps> {
     var {
       description,
       titleText,
       listName,
       isChecked,
       context,
       isConfigured,
       fields,
       dateSigned,
       siteName,
       checkboxLabel,
       CheckboxPlaceholder
     } = this.props;

     
     this._services = new SPService();

     return !this.state.checked && !isConfigured || !this.state.checked && isConfigured ? (
       <div className={styles.outerDiv} >
         <h3 className={`${styles.title}`}>{titleText}</h3>
         <div className={`${styles.descDiv}`}>
         <p>{description}</p>
        </div>
         <div className={`${styles.checkBoxDiv}`}>
           <Checkbox
             label={checkboxLabel}
             onChange={this._setCheckBoxValue.bind(this) }
            />
         </div>
       </div>
     ) : (
      <div className={styles.outerDiv} id="successLoad" >
      <Stack>
          {<SuccessExample signed={dateSigned === undefined || null ? this.state.date : dateSigned} listName={titleText}/>}
      </Stack>
      <div>
        <Checkbox
          className={`${styles.checkBoxDivDisabled}`}
          label={checkboxLabel}

          disabled
          defaultChecked
         />
      </div>
    </div>
     );
   }
   public getListName(listId:string){  
    const listName = this.props.fields.filter(item => item.key ==listId ).map( x => {
          return x.text;
     });  
     return listName[0]
  }

     public async _setCheckBoxValue(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,newValue: boolean) {


      const currentUserEmail = this.props.context.pageContext.user.email;

      const selectedSite = await this._services.getListsForSelectedSites(this.props.context, this.props.siteCollection)
   
      const selectedList: any = await selectedSite.value.filter((x:any) => x.Id ==this.props.listName).map((listname: any) => {return listname.Title});
      
      const itemObj = await this._services.getItemsFromSelectedList(this.props.context, selectedList[0],this.props.siteCollection)
      
      const items = await itemObj.res;
   
     if (newValue == true && selectedList[0] !== "") {
      
       const items: any = {
         currentUser: this.props.context.pageContext.user.email,
         context: this.props.context,
         siteCollection: this.props.siteCollection,
         listName: selectedList[0],
         Status: newValue,
         _spHttpContext: this.props.context.spHttpClient,
       };
       
       const response = await this._services.patchItemToSharePoint(items).catch((err) => {
        console.log(
          `Failed to post to Sharepoint list: ${selectedList[0]}. Error:`,err
        );
      });

      if(response !== undefined){
        if (response.response.ok && newValue) {
          this.isChecked  = newValue;
          const date = new Date(response.modified)
          this.setState({checked:response.checked, date:date.toDateString()});
          console.log("Success");
        } 
      }else {
        this.isChecked = false;
        this.setState({checked: this.isChecked})
        console.error("Couldn't find you in the selected list");     
      }
    }     
  }
}
