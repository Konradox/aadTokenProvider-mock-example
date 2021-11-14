import { WebPartContext } from "@microsoft/sp-webpart-base";
import { Log } from "@microsoft/sp-core-library";

export class TokenProvider {
	constructor(private context: WebPartContext) {}

	public async getToken(resource: string): Promise<string> {
		try {
			const tokenProvider = await this.context.aadTokenProviderFactory.getTokenProvider();
			const token = await tokenProvider.getToken(resource);
			return token;
		} catch (error) {
			Log.error("TokenProvider", error);
			throw error;
		}
	}
}
