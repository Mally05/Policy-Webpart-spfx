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
    You accepted this {p.listName} on <strong>{p.signed}</strong>
  </MessageBar>
);

 export default class Policy extends React.Component<IPolicyWebPartProps, CheckedState> {
  
   private _services: SPService;
   private isChecked: boolean = false;
   private listName: string;
   private dateSigned: any;

   constructor(props) {
    super(props);
    this.state = { checked: this.props.isChecked,
      date:""
    }
  }
  
   
   
componentDidUpdate(prevProps: Readonly<IPolicyWebPartProps>, prevState: Readonly<CheckedState>, snapshot?: any): void {
  if(prevState.checked !== this.isChecked){
    const listName = this.getListName(prevProps.listName);
    this._services.hasApprovedPolicyForSelectedList(this.props.context, listName)
      .then((res:any)=>{
        this.setState({checked: res.checked, date: res.modified})
        console.log(res,"ResUpdate")
      })
  }
    console.log(prevState.checked,"DidUpdateState")
}
componentDidMount(): void {
  this._services.loadDropdownListOptions(this.props.context).then((item:any) =>{
      
      item.filter((x:any) => x.key ==this.props.listName ).map(listname => {
      listname.text
      this._services.hasApprovedPolicyForSelectedList(this.props.context, listname.text)
     .then((res:any)=>{
       if(res.checked){
         this.setState({checked: res.checked, date: res.modified})
       } 2     
       console.log(res,"DidMount")
     })
      })
    });
     console.log("Didmount")
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
       dateSigned
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
             label="I accept the terms and conditions"
             onChange={this._setCheckBoxValue.bind(this) }
            />
         </div>
       </div>
     ) : (
      <div className={styles.outerDiv} >
      <Stack>
          {<SuccessExample signed={dateSigned === undefined || null ? this.state.date : dateSigned} listName={titleText}/>}
      </Stack>
      <div>
        <Checkbox
          className={`${styles.checkBoxDivDisabled}`}
          label="I accept the terms and conditions"
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

     this._services.loadDropdownListOptions(this.props.context).then((item:any) =>{
      
      item.filter((x:any) => x.key ==this.props.listName ).map(listname => {
      listname.text
      this._services.hasApprovedPolicyForSelectedList(this.props.context, listname.text);
     if (newValue == true && listname.text !== "") {
       const items: any = {
         currentUser: this.props.context.pageContext.user.email,
         context: this.props.context,
         listName: listname.text,
         Status: newValue,
         _spHttpContext: this.props.context.spHttpClient,
       };
        this._services.patchItemToSharePoint(items).then((response:any) => { 
           if (response.response.ok) {
            if(this.isChecked){
              this.isChecked  = true;
              const date = new Date(response.modified)
              this.setState({checked:response.checked, date:date.toDateString()});
              console.log("Success");
            }
           } else {
             this.isChecked = false;
             console.log("Failed to post");     
           }
         })
         .catch((err) => {
           console.log(
             `Failed to post to Sharepoint list: ${this.props.listName}. Error:`,err
           );
         });
     }     
      this._services.hasApprovedPolicyForSelectedList(this.props.context, listname.text);
      })})
    this.dateSigned = this.props.dateSigned;

     this.render();

     if(newValue == true){
      this.isChecked = newValue;
    }else{
      this.isChecked = newValue;
    }
     console.log("checked: ", this.state.checked);
   }
 }
