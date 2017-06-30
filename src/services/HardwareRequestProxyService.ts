
import { IHardwareRequest } from "../model/IHardwareRequest";
import { HttpClient, HttpClientConfiguration, IHttpClientOptions, HttpClientResponse } from "@microsoft/sp-http";

export const AzureFunctionUrl = "https://yp-labs-func.azurewebsites.net/api/AddHardwareRequest";
export const AzureFunctionSiteUrl = "https://yp-labs-func.azurewebsites.net";

export class HardwareRequestProxyService {

    constructor(private httpClient: HttpClient) {

    }

    public submitRequest(request: IHardwareRequest): Promise<HttpClientResponse> {
        return this.httpClient.post(AzureFunctionUrl, HttpClient.configurations.v1, {
            credentials: "include",
            mode: "cors",
            body: JSON.stringify(request)
        });
    }
}