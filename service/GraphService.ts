import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { IGraphService } from "./IGraphService";
import { MSGraphClientFactory, MSGraphClient } from "@microsoft/sp-http";

export class GraphService implements IGraphService {
	public static readonly serviceKey: ServiceKey<IGraphService> = ServiceKey.create<IGraphService>(
		"JT:JtBotGraph:GraphService",
		GraphService
	);

	private _gcf: MSGraphClientFactory;
	private _client: MSGraphClient;
	private get _gc() {
		return this._gcf.getClient("3");
	}

	public constructor(serviceScope: ServiceScope) {
		serviceScope.whenFinished(() => {
			this._gcf = serviceScope.consume(MSGraphClientFactory.serviceKey);
		});
	}
}
