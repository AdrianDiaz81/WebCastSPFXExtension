import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  IListViewCommandSetRefreshEventParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { HttpClient, SPHttpClient, HttpClientConfiguration, HttpClientResponse, ODataVersion, IHttpClientConfiguration, IHttpClientOptions, ISPHttpClientOptions } from '@microsoft/sp-http';

import * as strings from 'functionCommandSetStrings';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IFunctionCommandSetCommandSetProperties {
  // This is an example; replace with your own property
  disabledCommandIds: string[];
}

const LOG_SOURCE: string = 'FunctionCommandSetCommandSet';

export default class FunctionCommandSetCommandSet
  extends BaseListViewCommandSet<IFunctionCommandSetCommandSetProperties> {

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized FunctionCommandSetCommandSet');
    return Promise.resolve<void>();
  }

  @override
  public onRefreshCommand(event: IListViewCommandSetRefreshEventParameters): void {
    event.visible = true; // assume true by default

    if (this.properties.disabledCommandIds) {
      if (this.properties.disabledCommandIds.indexOf(event.commandId) >= 0) {
        Log.info(LOG_SOURCE, 'Hiding command ' + event.commandId);
        event.visible = false;
      }
    }
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {

    switch (event.commandId) {
      case 'COMMAND_1':
        alert(`Clicked ${strings.Command1}`);
        break;
      case 'COMMAND_2':
        alert(`Clicked ${strings.Command2}`);
        break;
      case 'COMMAND_3':
        this.Execute(event);
        break;
      default:
        throw new Error('Unknown command');
    }
  }

   Execute(id:any):void{
     alert("LLamo a la Azure Function");
        var siteUrl: string = this.context.pageContext.web.absoluteUrl;;
        var pageText:string="1";
        var pageName:string="Document"
     const functionUrl: string = "https://cob-pnp-functions.azurewebsites.net/api/CreateModernPage?code=WniDsXQ43Nf1HYB0JEIRuRrbLPaTTQnuithMnqtXoLQ54Hz6FY/j3g==";
    const requestHeaders: Headers = new Headers();
        requestHeaders.append("Content-type", "application/json");
        requestHeaders.append("Cache-Control", "no-cache");
    const postOptions: IHttpClientOptions = {
        headers: requestHeaders,
        body: `{ SiteUrl: '${siteUrl}', List: '${pageName}', Id: '${pageText}' }`
      };
          this.context.httpClient.post(functionUrl, HttpClient.configurations.v1, postOptions).then((response: HttpClientResponse) => {
         response.json().then((responseJSON: JSON) => {
           var responseText = JSON.stringify(responseJSON);
            alert(responseText)
          })
          .catch ((response: any) => {
            let errMsg: string = `WARNING - error when calling URL ${functionUrl}. Error = ${response.message}`;            
            console.log(errMsg);
            alert(errMsg);
          });
      });
  }
}
