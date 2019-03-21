import { ServiceKey, ServiceScope } from '@microsoft/sp-core-library';
import { AadHttpClientFactory, AadHttpClient } from '@microsoft/sp-http';

export interface ICustomService {
    executeMyRequest(): void;
}

export class CustomService implements ICustomService {

    public static readonly serviceKey: ServiceKey<ICustomService> =
        ServiceKey.create<ICustomService>('vrd:ICustomService', CustomService);
    
    private _aadHttpClientFactory: AadHttpClientFactory;
    private _aadHttpClient: AadHttpClient;

    constructor(serviceScope: ServiceScope) {
        serviceScope.whenFinished(() => {

            this._aadHttpClientFactory = serviceScope.consume(AadHttpClientFactory.serviceKey)

            this._aadHttpClientFactory.getClient("https://tenant.onmicrosoft.com/6b347c27-f360-47ac-b4d4-af78d0da4223").then((client: AadHttpClient) => {
                this._aadHttpClient = client;
            });
        });
    }

    public executeMyRequest(): void {
        this._aadHttpClient.get('https://myfunction.azurewebsites.net/api/CurrentUser', AadHttpClient.configurations.v1);
    }

}