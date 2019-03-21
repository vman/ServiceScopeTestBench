import { ServiceKey, ServiceScope } from '@microsoft/sp-core-library';
import { SPHttpClient } from '@microsoft/sp-http';

export interface ICustomGraphService {
    executeMyRequest(): void;
}

export class CustomGraphService implements ICustomGraphService {

    public static readonly serviceKey: ServiceKey<ICustomGraphService> =
        ServiceKey.create<ICustomGraphService>('vrd:ICustomGraphService', CustomGraphService);
    
    private _spHttpClient: SPHttpClient;

    constructor(serviceScope: ServiceScope) {
        serviceScope.whenFinished(() => {
            this._spHttpClient = serviceScope.consume(SPHttpClient.serviceKey)
        });
    }

    public executeMyRequest(): void {
        this._spHttpClient.get("", SPHttpClient.configurations.v1);
    }

}