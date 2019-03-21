import { ServiceKey, ServiceScope } from '@microsoft/sp-core-library';
import { MSGraphClientFactory, MSGraphClient } from '@microsoft/sp-http';

export interface ICustomGraphService {
    executeMyRequest(): void;
}

export class CustomGraphService implements ICustomGraphService {

    public static readonly serviceKey: ServiceKey<ICustomGraphService> =
        ServiceKey.create<ICustomGraphService>('vrd:ICustomGraphService', CustomGraphService);
    
    private _msGraphClientFactory: MSGraphClientFactory;
    private _msGraphClient: MSGraphClient;

    constructor(serviceScope: ServiceScope) {
        serviceScope.whenFinished(() => {

            this._msGraphClientFactory = serviceScope.consume(MSGraphClientFactory.serviceKey)

            this._msGraphClientFactory.getClient().then((client: MSGraphClient) => {
                this._msGraphClient = client;

                this._msGraphClient.api('/me').get((error, user: any, rawResponse?: any) => {
                    console.log(user);
                });;
            });
        });
    }

    public executeMyRequest(): void {
       // this._msGraphClient.api('/me');
    }

}