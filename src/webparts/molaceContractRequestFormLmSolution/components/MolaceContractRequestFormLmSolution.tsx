import * as React from 'react';
import { IMolaceContractRequestFormLmSolutionProps } from './IMolaceContractRequestFormLmSolutionProps';
import { IMolaceContractRequestFormLmSolutionState } from './IMolaceContractRequestFormLmSolutionState';
import styles from "./Trpreqfrm.module.scss"
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "@pnp/sp/site-users/web";
import "@pnp/sp/items";
import { TextField } from '@fluentui/react/lib/TextField';
import { PrimaryButton } from '@fluentui/react';
import {  MessageBar, MessageBarType, Spinner, SpinnerSize, Stack, getId } from 'office-ui-fabric-react';
import * as ReactDOM from 'react-dom';
import "@pnp/sp/folders";
import * as formconst from "../../constant";
import {  getCustomerTitle, updateData } from '../../../services/formservices';
import { isEmpty } from '@microsoft/sp-lodash-subset';
import { SPFI } from '@pnp/sp';
import { getSP } from '../../../pnpjsconfig';

 

  let isbuttondisbled : boolean = false;
  let buttontext : string = "Submit";
  let isLoading : boolean = false;
  let reqredirect : string
  let itemId:any
  let custRefcode : string
  let customerTitle:string
  let folderUrl: string;

export default class Trpreqfrm extends React.Component<IMolaceContractRequestFormLmSolutionProps, IMolaceContractRequestFormLmSolutionState> {


  private vdt: DataTransfer; 
  private odt: DataTransfer; 
  
 
  filesNamesRef: React.RefObject<HTMLSpanElement>;
  handleChange: any;
  pickerId: string;
  
  constructor(props: IMolaceContractRequestFormLmSolutionProps, state: IMolaceContractRequestFormLmSolutionState) {  
    super(props);  
   
    this.vdt = new DataTransfer();
    this.odt = new DataTransfer();
    this.filesNamesRef = React.createRef();
    this.pickerId = getId('inline-picker');
    this.state = {  
     
      title:"",
      voyage:"",
      addothers:"",
      isSuccess: false,
      files: [],
      vdocuments:"",
      odocuments:"",
      baf:"",
      onload:true,
      allfieldsvalid:true,
     
    }; 
    
  }

  public componentDidMount(){

  const url = document.location.search //window.location.href
  const urlParams = new URLSearchParams (url);
  const paramtitle = urlParams.get("Title");
  console.log(paramtitle)
  reqredirect = formconst.REQ_REDIRECT+"?"+"Title="+paramtitle
  itemId = paramtitle.match(/^(\d+)-/)[1];
  if (itemId.charAt(0) === "0") {itemId = itemId.substring(1);}
  console.log(itemId)
  custRefcode = paramtitle.match(/-(\w+)-/)[1];
  console.log(custRefcode)
  this.fetchCustomerItems();
  this.setState({title:paramtitle})
  }

  fetchCustomerItems = async () => {
    try {
      getCustomerTitle(this.props,custRefcode).then((customertitle: string)=>{

        customerTitle= customertitle
        console.log(customerTitle);
        
    
    });
    } catch (error) {
      console.error('Error fetching customer items:', error);
    }
  };
  


  vhandleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
   
  const filesNames = document.querySelector<HTMLSpanElement>('#vfilesList > #vfiles-names');
    for (let i = 0; i < e.target.files.length; i++) {
      let fileBloc = (
        <span key={i} className="file-block">
        <span className="name">{e.target.files.item(i).name}</span>
  <span className="file-delete">
    <button> Remove</button>
  </span>
  <br/>
        </span>
      );
      
      if (filesNames) {
        const fileBlocContainer = document.createElement('div');
        ReactDOM.render(fileBloc, fileBlocContainer);
        filesNames?.appendChild(fileBlocContainer.firstChild);
      }
    }
  
    for (let file of e.target.files as any) {
      this.vdt.items.add(file);
    }
  
    e.target.files = this.vdt.files;
  
    document.querySelectorAll('span.file-delete').forEach((button) => {
      button.addEventListener('click', () => {
        let name = button.nextSibling.textContent;
  
        (button.parentNode as HTMLElement)?.remove();
  
        for (let i = 0; i < this.vdt.items.length; i++) {
          if (name === this.vdt.items[i].getAsFile()?.name) {
            this.vdt.items.remove(i);
            continue;
          }
        }
  
        e.target.files = this.vdt.files;
    
      });
    });

  };
  
  ohandleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
   
  const filesNames = document.querySelector<HTMLSpanElement>('#ofilesList > #ofiles-names');
    for (let i = 0; i < e.target.files.length; i++) {
      let fileBloc = (
        <span key={i} className={'file-block'}>
         <span className="name">{e.target.files.item(i).name}</span>
          <span className="file-delete">
          <button> Remove</button>
         </span>
           <br/>
        </span>
      );
      
      if (filesNames) {
        const fileBlocContainer = document.createElement('div');
        ReactDOM.render(fileBloc, fileBlocContainer);
        filesNames?.appendChild(fileBlocContainer.firstChild);
   
      }
    }
  
    for (let file of e.target.files as any) {
      this.odt.items.add(file);
    }
  
    e.target.files = this.odt.files;
  
    document.querySelectorAll('span.file-delete').forEach((button) => {
      button.addEventListener('click', () => {
        let name = button.nextSibling.textContent;
  
        (button.parentNode as HTMLElement)?.remove();
  
        for (let i = 0; i < this.odt.items.length; i++) {
          if (name === this.odt.items[i].getAsFile()?.name) {
            this.odt.items.remove(i);
            continue;
          }
        }
  
        e.target.files = this.odt.files;
      });
    });
  };


  _updateItem  =async (props:IMolaceContractRequestFormLmSolutionProps):Promise<void>=>{
    
    
    folderUrl = formconst.LIBRARYNAME +"/" + customerTitle + "/"+ this.state.title;

     if(isEmpty(this.state.baf) || isEmpty(this.state.addothers))
     {
     this.setState({allfieldsvalid:false}) ; 
     console.log(this.state.allfieldsvalid)
    
     return;
     }else {
      this.setState({allfieldsvalid:true}) ; 
      isbuttondisbled = true;
      buttontext = "Saving...";
      isLoading = true;
     }
     
    // Update the item
    const updatedData = {
      VoyagePLContribution: this.state.voyage,
      Others: this.state.addothers,
      //VoyagePLContributionSupportingDo:this.state.vdocuments,
      //OthersSupportingDocuments:this.state.odocuments,
      BAF:this.state.baf
    };
    updateData(this.props,itemId, updatedData).then(async () => {
      await upload();
      const updatedDatadoclink = {
        //VoyagePLContribution: this.state.voyage,
       // Others: this.state.addothers,
        VoyagePLContributionSupportingDo:this.state.vdocuments,
        OthersSupportingDocuments:this.state.odocuments,
       // BAF:this.state.baf
      };
      updateData(this.props,itemId, updatedDatadoclink).then(async () => {
      isbuttondisbled = false;
      buttontext = "Submit";
      isLoading = false;
      this.setState({ isSuccess: true });
      window.open(formconst.SUBMIT_REDIRECT,"_self")
      })
    }).catch((error: any) => {
    
    var obj = JSON.stringify(error);
  
    if(obj.indexOf("400") !== -1)
    {    console.log("mATCH FOUND")
         
         this.setState({allfieldsvalid:false}) 
    }else{
    console.log('Error:', error);}
  });

  const upload = async () => {
  
    console.log(folderUrl)
    
    let vstrbgurl = "";
    let ostrbgurl = "";
    const _sp :SPFI = getSP(props.context) ;
    //_sp.web.folders.addUsingPath(folderUrl);
    // vfiles
    let vfileurl = [];
    let vinput = document.getElementById("vattachment") as HTMLInputElement;
    const vcategory = 'Voyage P/L Contribution'
    console.log(vinput.files);
    if (vinput.files.length > 0) {
      let vfiles = vinput.files;
      
      for (var i = 0; i < vfiles.length; i++) {
        let vfile = vinput.files[i];
        console.log("vfile",vfile)
        vfileurl.push(formconst.WEB_URL + "/" + folderUrl + "/" + vfile.name);
        try {
          let vuploadedFile = await _sp.web.getFolderByServerRelativePath(folderUrl).files.addChunked(vfile.name, vfile, (data) => {
            console.log("File uploaded successfully");
          });
          let item = await vuploadedFile.file.getItem();
          await item.update({Section:vcategory});
          await item.update({RequestID:this.state.title})
        } catch (err) {
          console.error("Error uploading file:", err);
        }
      }
      let vconvertedStr = vfileurl.map(url => `<a href="${url.trim()}" target="_blank">${url.substring(url.lastIndexOf("/") + 1)}</a>`);
     vstrbgurl = vconvertedStr.toString();
    //console.log(vstrbgurl);
    this.setState({ vdocuments: vstrbgurl });
    
    } else {
      console.log("No file selected for upload.");
      
    }
    
 
    let ofileurl = [];
    let oinput = document.getElementById("othersattachment") as HTMLInputElement;
    const ocategory = 'Others'
    console.log(oinput.files);
   
    if (oinput.files.length > 0) {
      let ofiles = oinput.files;
   
      for (var i = 0; i < ofiles.length; i++) {
        let ofile = oinput.files[i];
        console.log("ofile",ofile)
        ofileurl.push(formconst.WEB_URL+ "/" + folderUrl + "/" + ofile.name);
        try {
          
          let ouploadedFile = await _sp.web.getFolderByServerRelativePath(folderUrl).files.addChunked(ofile.name, ofile, (data) => {
            console.log("File uploaded successfully");
          });
          let item = await ouploadedFile.file.getItem();
          await item.update({Section:ocategory});
          await item.update({RequestID:this.state.title}) ;
             
          
        } catch (err) {
          console.error("Error uploading file:", err);
        }
      }
      let oconvertedStr = ofileurl.map(url => `<a href="${url.trim()}" target="_blank">${url.substring(url.lastIndexOf("/") + 1)}</a>`);
       ostrbgurl = oconvertedStr.toString();
      //console.log(ostrbgurl);
      this.setState({ odocuments: ostrbgurl });
      
    } else {
      console.log("No file selected for upload.");
      
    }

  }
  }

handleInputChange(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, field: string) {
  const { value } = event.target as HTMLInputElement;

  // Update the state based on the field parameter
  if (field === 'voyage') {
    this.setState({ voyage: value });
  } else if (field === 'baf') {
    this.setState({ baf: value });
  } else if (field === 'others') {
    this.setState({ addothers: value });
  }
}

cancel =()=>{
  window.open(formconst.CANCEL_REDIRECT,"_self");
}
  public render(): React.ReactElement<IMolaceContractRequestFormLmSolutionProps> {
 
    let successMessage : JSX.Element | null
 
    let bafFieldErrorMessage: JSX.Element | null
    let othersFieldErrorMessage: JSX.Element | null
  
    let EmailFieldErrorMessage: JSX.Element | null
    let FormFieldErrorMessage: JSX.Element | null

    successMessage = this.state.isSuccess ?
    <MessageBar messageBarType={MessageBarType.success}>Request Id : {this.state.title} submitted successfully.</MessageBar>
    : null;
    
 

    if(!this.state.allfieldsvalid){
   
 
        bafFieldErrorMessage = isEmpty(this.state.baf) ?
        <MessageBar messageBarType={MessageBarType.error}>BAF is required.</MessageBar>
        : null;
        othersFieldErrorMessage = isEmpty(this.state.addothers) ?
        <MessageBar messageBarType={MessageBarType.error}>Others is required.</MessageBar>
        : null;
  
        FormFieldErrorMessage= 
        <MessageBar messageBarType={MessageBarType.error}>Please provide all required information and submit the form.</MessageBar>
        
      }
    return (
    
    <section>
      <div>
        <p>As a line manager, you have been requested to provide additiponal information as part of a new transportation contract request.</p>
        <p><b>Request No: </b><a href = {reqredirect} target='_blank'>{this.state.title}</a></p>
      </div>
        <div>
          <p className={styles.heading}>Additional Information</p>
          <div>
          <p><span className={styles.required}><b>*</b></span> Required.</p>
    </div>
        
      <p className={styles.formlabel}>Voyage P/L Contribution</p>
      <TextField value={this.state.voyage} onChange={(event) => this.handleInputChange(event, 'voyage')} multiline rows={3} />
      <div className="mt-5 text-center">
        <label htmlFor="vattachment" className="btn btn-primary text-light" role="button" aria-disabled="false">
          + Add Supporting Documents
        </label>
        <input
          type="file"
          name="file[]"
          accept=""
          id="vattachment"
          style={{ visibility: 'hidden', position: 'absolute' }}
          multiple
          onChange={this.vhandleFileUpload}
        />
      <p id="vfiles-area">
          <span id="vfilesList">
            <span ref={this.filesNamesRef} id="vfiles-names"></span>
          </span>
        </p>
      </div>
      
      <p className={styles.formlabel}>BAF<span className={styles.required}> *</span></p>  
      <TextField value={this.state.baf} onChange={(event) => this.handleInputChange(event, 'baf')}/>{bafFieldErrorMessage}
    
      <p className={styles.formlabel}>Others<span className={styles.required}> *</span></p>
      <TextField value={this.state.addothers} onChange={(event) => this.handleInputChange(event, 'others')} multiline rows={3} />{othersFieldErrorMessage}
      <div className="mt-5 text-center">
        <label htmlFor="othersattachment" className="btn btn-primary text-light" role="button" aria-disabled="false">
          + Add Supporting Documents
        </label>
        <input
          type="file"
          name="file[]"
          accept=""
          id="othersattachment"
          style={{ visibility: 'hidden', position: 'absolute' }}
          multiple
          onChange={this.ohandleFileUpload}
        />
      <p id="ofiles-area">
          <span id="ofilesList">
            <span ref={this.filesNamesRef} id="ofiles-names"></span>
          </span>
        </p>
      </div>
        <br/>
        {EmailFieldErrorMessage}
    <Stack horizontal horizontalAlign='end' className={styles.stackContainer}>     
    <PrimaryButton text={buttontext} onClick={() => this._updateItem(this.props)} disabled= {isbuttondisbled}/>
    <PrimaryButton text="Cancel"  onClick ={this.cancel}/>
   
    </Stack> 
    <br />
    <div>      
      {isLoading && <Spinner label="Saving, please wait..." size={SpinnerSize.large} />}
      </div>
   
    <br />
    {FormFieldErrorMessage}
    {successMessage}
    
        </div>
      </section>
    );
  }
}

