import { Client } from '@microsoft/microsoft-graph-client';
import { User } from '@microsoft/microsoft-graph-types';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import SharePointAuthenticationProvider from './SharePointAuthenticationProvider';

export interface IHelloWorldWebPartProps {
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {
  private client: Client = null;

  public async render(): Promise<void> {
    const me = await this.client.api('/me').get() as User;
    this.domElement.innerHTML = `
      <h1>
        Hello ${me.displayName}!
      </h1>
    `;
  }

  protected async onInit(): Promise<void> {
    this.client = Client.initWithMiddleware({
      authProvider: new SharePointAuthenticationProvider(this.context),
    });

    let c = await this.context.msGraphClientFactory.getClient('3');

    return super.onInit();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
}
