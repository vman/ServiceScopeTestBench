import { ServiceKey, ServiceScope } from '@microsoft/sp-core-library';
import { AadHttpClientFactory, AadHttpClient, HttpClientResponse } from '@microsoft/sp-http';

export interface ICustomService {
    executeMyRequest(): Promise<JSON>;
}

export class CustomService implements ICustomService {

    public static readonly serviceKey: ServiceKey<ICustomService> =
        ServiceKey.create<ICustomService>('my-custom-app:ICustomService', CustomService);

    private _aadHttpClientFactory: AadHttpClientFactory;

    constructor(serviceScope: ServiceScope) {
        serviceScope.whenFinished(() => {
            this._aadHttpClientFactory = serviceScope.consume(AadHttpClientFactory.serviceKey);
        });
    }

    public executeMyRequest(): Promise<JSON> {
        //You can add your own AAD resource here. Using the Graph API resource for simplicity.
        return this._aadHttpClientFactory.getClient("https://graph.microsoft.com").then((_aadHttpClient: AadHttpClient) => {
            return _aadHttpClient.get('https://graph.microsoft.com/v1.0/me', AadHttpClient.configurations.v1).then((response: HttpClientResponse) => {
                return response.json();
            });
        });
    }

}