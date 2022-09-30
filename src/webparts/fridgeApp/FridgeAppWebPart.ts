import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
PropertyPaneDropdown,
PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './FridgeAppWebPart.module.scss';
import * as strings from 'FridgeAppWebPartStrings';

import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions
} from '@microsoft/sp-http';
import { AppLocalEnvironmentTeams } from 'FridgeAppWebPartStrings';

//Test Variables
export interface IFridgeAppWebPartProps {
  description: string;
  test: string;
  test1: boolean;
  test2: string;
  test3: boolean;
}

//List and list item for the config list
export interface ISPLists {
  value: ISPList[];
}
export interface ISPList {
  Title: string; 
  NumberofItem: number;
  Id: string;
}


export default class FridgeAppWebPart extends BaseClientSideWebPart<IFridgeAppWebPartProps> {

  //Config for environment
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private currentListVal: ISPList[];
  testValue: string ="";

  //Base rendering method
  public render(): void {
    this.domElement.innerHTML = `
<section class="${styles.fridgeApp} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
  <div class="${styles.welcome}">
    <img alt="" src="${this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png')}" class="${styles.welcomeImage}" />
    <h2>Welcome to The Fridge, ${escape(this.context.pageContext.user.displayName)}!</h2>
  </div>
  <div>
    <div>Web part description: <strong>${escape(this.properties.description)}</strong></div>
    <div>Loading from: <strong>${escape(this.context.pageContext.web.title)}</strong></div>
    <div>List: <strong>${escape(this.testValue)}</strong>
    <div>Search for Item: <input type="text" id="itemSearchName" name= "itemSearchName"><input type="button" id="searchQueryBttn" value="Search"> </div>
    <div><label for="newItemName">Input new item:</label> <input type="text" id="newItemName" name="newItemName"><input type="number" id="newItemNumber" name="newItemNumber"> <input type="button" id="newItemBttn" value="Add New Item"></div>
    <div><label for="existingItemName">Edit existing item:</label><input type="text" id="existingItemName" name="existingItemName"><input type="number" id="existingItemNumber" name="existingItemNumber"><input type="button" id="existingItemBttn"value="Alter item"></div>
  </div>
  <div id="spListContainer" />
</section>`;

//binds for clicks and the rendering function for the Sharepoint list
this._renderListAsync();
this._bindSave();
this._bindSearch();
this._bindUpdate();
  }

//bind methods
private _bindSave(): void {
  this.domElement.querySelector('#newItemBttn').addEventListener('click',() =>{this.addNewItem();setTimeout(() => { this._bindUpdateItem(); }, 750)});
  
}
private _bindSearch():void {
  this.domElement.querySelector('#searchQueryBttn').addEventListener('click',()=>{this.searchForItem()});
}
private _bindUpdateItem(): void {
  const collection = document.getElementsByClassName("editBttnClass");
  const webPart: FridgeAppWebPart = this; 
  let buttons = this.domElement.getElementsByClassName('editBttnClass');
  for(let i= 0; i < buttons.length;i++){
      
      if(buttons[i].id != null){
      
      document.getElementById(buttons[i].id).addEventListener('click',() => {this.renderItemView(buttons[i]);});
      }
    };
}
private _bindUpdate():void{ 
  this.domElement.querySelector('#existingItemBttn').addEventListener('click',()=>this.updateExistingItem(this.currentListVal));
}

//Method for adding a new fridge item
private addNewItem(): void {
  const varName = (<HTMLInputElement>document.getElementById("newItemName")).value;
  const varNumber = (<HTMLInputElement>document.getElementById("newItemNumber")).value;
  let duplicatefound = false;
  for(let i = 0; i < this.currentListVal.length;i++){
if(this.currentListVal[i].Title == varName){
  duplicatefound = true;
}
  }
  if(duplicatefound == false){

  const siteURL: string = this.context.pageContext.site.absoluteUrl+ "/_api/web/lists/getbytitle('FridgeConfig')/items";
  this.testValue = siteURL;
  const itemBody: any = {
    "Title": varName,
    "NumberofItem": varNumber
  };
  const spHttpClientOptions: ISPHttpClientOptions = {
    "body": JSON.stringify(itemBody)
  };

  this.context.spHttpClient.post(siteURL,SPHttpClient.configurations.v1,spHttpClientOptions).then((response: SPHttpClientResponse) => {if(response.ok){
    this._renderListAsync();
  }});
}
else{
  alert(varName + " is already in the fridge!")
}
}



private searchForItem():void{
let items = this.currentListVal;
let newlist: ISPList[];
newlist = [];
let newlistItem: ISPList;

const searchQuery = (<HTMLInputElement>document.getElementById("itemSearchName")).value;
if(searchQuery == "" || searchQuery == null){
  this._renderListAsync();
}
for(let i = 0; i< items.length;i++){
  newlistItem = {Title:"",Id:"",NumberofItem:null}
if(searchQuery == items[i].Title){
  newlistItem={
    Title: items[i].Title,
    Id: items[i].Id,
    NumberofItem: items[i].NumberofItem
  };

  newlist.push(newlistItem);
}
}
this._renderList(newlist);
}


private updateExistingItem(items: ISPList[]):void{
  const varName = (<HTMLInputElement>document.getElementById("existingItemName")).value;
  const varNumber = (<HTMLInputElement>document.getElementById("existingItemNumber")).value;
  const headers: any = {
    "X-HTTP-Method": "MERGE", 
    "IF-MATCH": "*"
  };
  var varID: string = ""
  items.forEach((item: ISPList) => {
if(item.Title == varName){
   varID = item.Id;
} 
  });
if(varID != ""){
  const siteURL: string = this.context.pageContext.site.absoluteUrl+ "/_api/web/lists/getbytitle('FridgeConfig')/items("+ varID +")";
  const itemBody: any = {
    "Title": varName,
    "NumberofItem": varNumber
  };
  const spHttpClientOptions: ISPHttpClientOptions = {
    "headers": headers,
    "body": JSON.stringify(itemBody)
  };
this.context.spHttpClient.post(siteURL,SPHttpClient.configurations.v1,spHttpClientOptions).then((response: SPHttpClientResponse) =>{if(response.ok){
  this._renderListAsync();
}});
}
else{
  alert(varName + " is not currently in the fridge");
}
}



public renderItemView(item: Element):void{
  const varName = (<HTMLInputElement>document.getElementById(item.id)).id;
  const headers: any = {
    "X-HTTP-Method": "DELETE", 
    "IF-MATCH": "*"
  };
  var varID: string = ""
  this.currentListVal.forEach((item: ISPList) => {
   let nameValue = varName.split("editbutton")
if(item.Title == nameValue[1]){
   varID = item.Id;
} 
  });
if(varID != ""){
  const siteURL: string = this.context.pageContext.site.absoluteUrl+ "/_api/web/lists/getbytitle('FridgeConfig')/items("+ varID +")";
  const itemBody: any = {
    
  };
  const spHttpClientOptions: ISPHttpClientOptions = {
    "headers": headers,
    "body": JSON.stringify(itemBody)
  };
this.context.spHttpClient.post(siteURL,SPHttpClient.configurations.v1,spHttpClientOptions).then((response: SPHttpClientResponse) =>{if(response.ok){
  this._renderListAsync();
}});
}
}


  private _renderListAsync(): void {
    this._getListData()
      .then((response) => {
        this.currentListVal = response.value;
        this._renderList(response.value);
      });
      setTimeout(() => { this._bindUpdateItem(); }, 750)
  }
  public testPress = (Title1: string) => {
    alert("Test " + Title1)
  }
//<ul class="${styles.list}">
//<li class="${styles.listItem}">
//<span class="ms-font-l">${item.Title}</span>
       // <div class = "ms-font-l">${item.NumberofItem}</div>
    //  </li>
   // </ul>
  private _renderList(items: ISPList[]): void {

    let html: string = `<table border=2 width=100% style="font-family: "Trebuchet MS", Arial, Helvetica, sans-serif; id="itemTable"}>`;
    html += '<b><th style="background-color: #3bc2ed;" >Item</th> <th style="background-color: #3bc2ed; width:1px;">Number of Item </th> <th style="background-color: #3bc2ed; width:1px;">Delete</th>'
    items.forEach((item: ISPList) => {
      var curItem = item.Title
      html += `
     
      <tr style="background-color: #c4e1e5;" id="tablerow${curItem}">             
      <td>${item.Title}</td>
      <td>${item.NumberofItem}</td>
      <td><button type="submit" 
      id="editbutton${item.Title}" class="editBttnClass" value="Edit" ">Delete</button>
      </td>
      </tr>
        `;
    });
  html += `</table>`

    const listContainer: Element = this.domElement.querySelector('#spListContainer');
    listContainer.innerHTML = html;
  }

  
  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();
this.testValue = "FridgeConfig";
    return super.onInit();
  }



  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
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
                label: 'Description'
              }),
              PropertyPaneTextField('test', {
                label: 'Multi-line Text Field',
                multiline: true
              }),
              PropertyPaneCheckbox('test1', {
                text: 'Checkbox'
              }),
              PropertyPaneDropdown('test2', {
                label: 'Dropdown',
                options: [
                  { key: '1', text: 'One' },
                  { key: '2', text: 'Two' },
                  { key: '3', text: 'Three' },
                  { key: '4', text: 'Four' }
                ]}),
              PropertyPaneToggle('test3', {
                label: 'Toggle',
                onText: 'On',
                offText: 'Off'
              })
            ]
            }
          ]
        }
      ]
    };
  }

  private _getListData(): Promise<ISPLists> {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists/getbytitle('FridgeConfig')/items`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        response.json;
        return response.json();
      });
  }



}

      function testCase(): EventListenerOrEventListenerObject {
        throw new Error('Function not implemented.');
      }


      function getbytitle(arg0: string) {
        throw new Error('Function not implemented.');
      }

function RemoveItem(_title: string) {
  alert("Item to be removed: " + _title)
}


