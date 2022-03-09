import { AuthenticationProvider } from "@microsoft/microsoft-graph-client";

/**
 * AadTokenProvider
 *
 * @interface AadTokenProvider
 */
declare interface AadTokenProvider {
    /**
     * get token with x
     *
     * @param {string} x
     * @memberof AadTokenProvider
     */
    getToken(x: string): Promise<string>;
}

/**
 * The instance of AadTokenProviderFactory created for this instance of component
 *
 * @export
 * @interface AadTokenProviderFactory
 */
 declare interface AadTokenProviderFactory {
    /**
     * Returns an instance of the AadTokenProvider that communicates with the current tenant's configurable
     * Service Principal.
     */
    getTokenProvider(): Promise<AadTokenProvider>;
}

/**
 * contains the contextual services available to a web part
 *
 * @export
 * @interface BaseComponentContext
 */
declare interface BaseComponentContext {
    /**
     * The instance of AadTokenProviderFactory created for this instance of component
     */
    aadTokenProviderFactory: AadTokenProviderFactory;
}

export default class SharePointAuthenticationProvider implements AuthenticationProvider {
    
    private context: BaseComponentContext = null;

    public constructor(context: BaseComponentContext) {
        this.context = context;
    }
    
    public async getAccessToken() : Promise<string> {
        const tokenProvider: AadTokenProvider = await this.context.aadTokenProviderFactory.getTokenProvider();
        return await tokenProvider.getToken('https://graph.microsoft.com');
    }
    
}