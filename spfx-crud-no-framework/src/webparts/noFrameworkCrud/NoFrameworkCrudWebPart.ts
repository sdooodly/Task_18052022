///importing the reqd 
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';
import styles from './NoFrameworkCrudWebPart.module.scss';
import * as strings from 'NoFrameworkCrudWebPartStrings';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http'; 
import { IListItem } from './loc/IListItem';
///Property pane config
export interface INoFrameworkCrudWebPartProps {
  listName: string;
  listField: string;
  description: string;
}
///Main render function
export default class NoFrameworkCrudWebPart extends BaseClientSideWebPart<INoFrameworkCrudWebPartProps> {
  public render(): void {
    ///HTML (Button, heading, etc)
    this.domElement.innerHTML = `
    <section class="${styles.noFrameworkCrud} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
        <h1><i>List name: ${escape(this.properties.listName)}</i></h1>
        <div><i>Purpose: Sample application to read data from a list.</i></div>     
      <div>
      <div class="${ styles.noFrameworkCrud }">                       
                 <button class="${styles.button} read-Button">  
                   <span class="${styles.label}"><i>Read the list</i></span>  
                 </button>               
            <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">  
              <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">  
                <div class="status"></div>  
                <ul class="items"><ul>  
              </div>  
            </div>    
    </div>`;  
    this.setButtonsEventHandlers();
  }
  ///Button config function
  private setButtonsEventHandlers(): void {  
    const webPart: NoFrameworkCrudWebPart = this;    
    this.domElement.querySelector('button.read-Button').addEventListener('click', () => { webPart.readItem(); });    
  } 
  ///Fuction reading from the list using REST API query
  private readItem(): void {  
    this.getLatestItemId()  
      .then((itemId: number): Promise<SPHttpClientResponse> => {  
        if (itemId === -1) {  
          throw new Error('The list is empty.');  
        }  
    
        this.updateStatus(`Loading information about item ID: ${itemId}...`);  
        ///To do : Check here to get other list field data  
        ///return this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.listName}')/items(${itemId})?$select=Title,Id`,
        return this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.listName}')/items(${itemId})?$select=Title,Id,Address,Number`,
          SPHttpClient.configurations.v1,  
          {  
            headers: {  
              'Accept': 'application/json;odata=nometadata',  
              'odata-version': ''  
            }  
          });  
      })  
      .then((response: SPHttpClientResponse): Promise<IListItem> => {  
        return response.json();  
      })  
      .then((item: IListItem): void => {  
        this.updateStatus(`ID: ${item.Id}, 
        Title: ${item.Title}, 
        Multiline data : ${item.Address},
        Number : ${item.Number}, 
        ${escape(this.properties.listField)} ??? `);
        ///this.updateStatus(`ID: ${item.Id}, Title: ${item.Title}, Multiline data : ${item.Address}, Number : ${item.Number},${escape(this.properties.listField)} ??? `);   
      }, (error: any): void => {  
        this.updateStatus('Loading latest item failed with error: ' + error);  
      });  
  }  
  ///Function to get ID of the last row in the list
  private getLatestItemId(): Promise<number> {  
    return new Promise<number>((resolve: (itemId: number) => void, reject: (error: any) => void): void => {  
      this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.listName}')/items?$orderby=Id desc&$top=1&$select=id`,  
        SPHttpClient.configurations.v1,  
        {  
          headers: {  
            'Accept': 'application/json;odata=nometadata',  
            'odata-version': ''  
          }  
        })  
        .then((response: SPHttpClientResponse): Promise<{ value: { Id: number }[] }> => {  
          return response.json();  
        }, (error: any): void => {  
          reject(error);  
        })  
        .then((response: { value: { Id: number }[] }): void => {  
          if (response.value.length === 0) {  
            resolve(-1);  
          }  
          else {  
            resolve(response.value[0].Id);  
          }  
        });  
    });  
  }
  ///Functions to update status
  private updateStatus(status: string, items: IListItem[] = []): void {  
    this.domElement.querySelector('.status').innerHTML = status;  
    this.updateItemsHtml(items);  
  }    
  private updateItemsHtml(items: IListItem[]): void {  
    ///this.domElement.querySelector('.items').innerHTML = items.map(item => `<li>${item.Title} (${item.Id}) ${item.Address} </li>`).join(""); 
     this.domElement.querySelector('.items').innerHTML = items.map(item => `<li>${item.Title} (${item.Id}) ${item.Address} ${item.Number} </li>`).join("");
  }
  ///Default  
  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }
    const {
      semanticColors
    } = currentTheme;
    this.domElement.style.setProperty('--bodyText', semanticColors.bodyText);
    this.domElement.style.setProperty('--link', semanticColors.link);
    this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered);
  }
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
  ///Property pane configuration
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'Enter the list name below.'
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('listName', {
                  label: strings.ListNameFieldLabel
                }),
                PropertyPaneTextField('listField', {
                  label:'Field',
                  }),

                PropertyPaneTextField('description', {
                  label:'Description',
                  multiline : true
                  })
               
              ]
            }
          ]
        }
      ]
    };
  }
    
}
