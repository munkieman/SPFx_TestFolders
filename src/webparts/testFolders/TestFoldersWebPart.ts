import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './TestFoldersWebPart.module.scss';
import * as strings from 'TestFoldersWebPartStrings';
import { SPComponentLoader } from '@microsoft/sp-loader';

//import {
  //SPHttpClient,
//  SPHttpClientResponse
//} from '@microsoft/sp-http';

import {spfi, SPFx} from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/items/get-all";
import { LogLevel, PnPLogging } from "@pnp/logging";

require('bootstrap');

export interface ITestFoldersWebPartProps {
  description: string;
  folderNameIDarray: any;
  dataResults: any[];
}

export default class TestFoldersWebPart extends BaseClientSideWebPart<ITestFoldersWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  //private asmDC = Web("https://munkieman.sharepoint.com/sites/asm_dc/"); 

  private _getListData(libraryName:string):Promise<any>{
    const sp = spfi().using(SPFx(this.context)).using(PnPLogging(LogLevel.Warning));
    //const asmlist = this.asmDC.lists.getByTitle(libraryName);

    const view : string = '<View><Query>' +
                            '<Where>' +
                                '<Eq>' +                
                                  '<FieldRef Name="Team"/>'+
                                  '<Value Type="TaxonomyFieldType">ASM Team A</Value>'+
                                '</Eq>' +
                            '</Where>'+
                            '<OrderBy>'+
                              '<FieldRef Name="Folder" Ascending="TRUE" />'+
                              '<FieldRef Name="SubFolder01" Ascending="TRUE" />'+
                              '<FieldRef Name="SubFolder02" Ascending="TRUE" />'+
                            '</OrderBy>'+
                          '</Query></View>';    

    return sp.web.lists.getByTitle('Policies').getItemsByCAMLQuery({ViewXml:view}, "FieldValuesAsText/FileRef", "FieldValueAsText/FileLeafRef")
      .then((response) => {
        if(response.length>0){
          console.log(response);
          return response;
        }else{
          return false;
        }
      }); 
      
    /*
    asmlist.getItemsByCAMLQuery({ViewXml:view}, "FieldValuesAsText/FileRef", "FieldValueAsText/FileLeafRef")
      .then((asm_Results) => {
        if(asm_Results.length>0){
          //for(let c=0;c<asm_Results.length;c++){
          //  this.properties.dataResults.push(asm_Results[c]);
          //}
          //this.properties.dataResults.push(asm_Results);
          console.log("ASM DC Results");
          console.log(asm_Results);  
          return asm_Results;
        }else{
          return false;
        }
      });
      return;
      */      
  }

  private _renderFolders(results:any[]): void{
    alert('getting folders');
    
    let html : string = "";
    let folderName : string = "";
    let folderNamePrev : string = "";
    let x=0;
    let count = 0;
    this.properties.folderNameIDarray=[];

    //console.log(results);

    results.forEach(()=>{
      console.log('folderName='+results[x].FieldValuesAsText.Folder);
      folderName = results[x].FieldValuesAsText.Folder;      

      if(folderName !== folderNamePrev){  
        let folderNameID=folderName.replace(/\s+/g, "")+x;
        html+=`<ul>
                <li>
                  <button class="btn btn-primary" id="${folderNameID}" type="button" data-bs-toggle="collapse" aria-expanded="true" aria-controls="accordionPF${x}">
                    <i class="bi bi-folder2"></i>
                    <a href="#" class="text-white ms-1">${folderName}</a>
                    <span class="badge bg-secondary">${count}</span>                    
                  </button>
                </li>
              </ul>`;            
        folderNamePrev=folderName;
        this.properties.folderNameIDarray.push(folderNameID,folderName);
        count++;
      }
      x++;
    });
    console.log(this.properties.folderNameIDarray);

    const listContainer: Element = this.domElement.querySelector('#folderContainer');
    listContainer.innerHTML = html;    

  }

  private _renderDataAsync(): void {

    this._getListData('Policies')
      .then((response) => {
        if(response.length>0){
          this._renderFolders(response);
        }else{
          alert("No Data found for this Team");
        }
      });
  }

  private folderListeners() : void {
    alert('adding folder listeners');

    for(let fn=0;fn<this.properties.folderNameIDarray.length;fn++){
      if(fn % 2 === 0){
        document.getElementById(this.properties.folderNameIDarray[fn]).addEventListener("click",(e:Event) => this.getFiles(this.properties.folderNameIDarray[fn+1]));
      }  
    };
  }

  private getFiles(folderName:string){
    alert(folderName);
  }

  public render(): void {
    const bootstrapCssURL = "https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css";
    const fontawesomeCssURL = "https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.11.2/css/regular.min.css";
    SPComponentLoader.loadCss(bootstrapCssURL);
    SPComponentLoader.loadCss(fontawesomeCssURL);

    this.domElement.innerHTML = `
    <section class="${styles.testFolders} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <div class="${styles.welcome}">
        <img alt="" src="${this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png')}" class="${styles.welcomeImage}" />
        <h2>Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>
        <div>${this._environmentMessage}</div>
        <div>Web part property value: <strong>${escape(this.properties.description)}</strong></div>
      </div>

      <div class="row">
        <h3>Libraries</h3>
      </div>
      <div class="row btnContainer btn-group">
        <button class="btn btn-primary" id="policies">Policies</button>
        <button class="btn btn-primary" id="procedures">Procedures</button>
        <button class="btn btn-primary" id="guides">Guides</button>
        <button class="btn btn-primary" id="forms">Forms</button>
        <button class="btn btn-primary" id="general">General</button>
        <button class="btn btn-primary" id="management">Management</button>
        <button class="btn btn-primary" id="archive">Archive</button>
      </div>
      <div class="row searchContainer">
        <input type="text" placeholder="search"/>
      </div>
      <div class="row">
        <div class="col-6" id="docFolders">
          <h4 class="colTitle">Folder</h4>
          <div class="justify-content-center flex-column colContainer" id="folderContainer"></div>
        </div>   
        <div class="col-6" id="docFiles">
          <h4 class="colTitle">Files</h4>
          <div class="justify-content-center flex-column colContainer" id="fileContainer"></div>
        </div>
      </div>
      <div id="spListContainer"/>
    </section>`;

    this._renderDataAsync();
    setTimeout(()=> {
      this.folderListeners();
    }
    ,2000);
  }

  protected async onInit(): Promise<void> {
    await super.onInit();

    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              throw new Error('Unknown host');
          }

          return environmentMessage;
        });
    }
    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
