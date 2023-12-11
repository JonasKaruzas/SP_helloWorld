import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './HelloWorldPartNameWebPart.module.scss';
import * as strings from 'HelloWorldPartNameWebPartStrings';

import { SPFx, spfi } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";





export interface IHelloWorldPartNameWebPartProps {
  description: string;
}

export default class HelloWorldPartNameWebPart extends BaseClientSideWebPart<IHelloWorldPartNameWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  
  
  public render(): void {

    this.domElement.innerHTML = `
    <section class="${styles.helloWorldPartName} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <div class="${styles.welcome}">
        <img alt="" src="${this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png')}" class="${styles.welcomeImage}" />
        <h2>Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>
        <div>${this._environmentMessage}</div>
      </div>
      <div>
        <h3>Welcome to SharePoint Framework!</h3>
        <div>Web part description: <strong>${escape(this.properties.description)}</strong></div>
        <div>Loading from: <strong>${escape(this.context.pageContext.web.title)}</strong></div>
        <div>${this.context.pageContext.list?.title}</div>
      </div>
    </section>`;
  }

private async test() {
  const context = this.context;
  const sp = spfi().using(SPFx(context));
  
  // create a new list, passing only the title
  const listAddResult = await sp.web.lists.add("My new list");
  
  // we can work with the list created using the IListAddResult.list property:
  const r = await listAddResult.list.select("Title")();
  
  // log newly created list title to console
  console.log(r.Title);

}

  protected async onInit(): Promise<void> {

    await this.test();




    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  // private async ensureCustomList(): Promise<void> {
  //   try {
  //     const listTitle = 'YourCustomList';

  //     // Check if the list already exists
  //     const listExists = await sp.web.lists.getByTitle(listTitle).get();

  //     // If the list doesn't exist, create it
  //     if (!listExists) {
  //       const listCreationResult: IListEnsureResult = await sp.web.lists.ensure(listTitle, 'Description of Your Custom List', 100);

  //       // Add other list settings and columns here if needed
  //       // Example: await listCreationResult.list.fields.addText('YourColumn', { maxLength: 255 });
  //     }
  //   } catch (error) {
  //     console.error('Error ensuring custom list:', error);
  //   }
  // }



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
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
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
