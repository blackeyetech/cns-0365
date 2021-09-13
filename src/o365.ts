// Imports here
import CNShell from "cn-shell";
import qs from "qs";

// Misc config consts here
const CFG_MS_GRAPH_APP_ID = "MS_GRAPH_APP_ID";
const CFG_MS_GRAPH_CLIENT_SECRET = "MS_GRAPH_CLIENT_SECRET";
const CFG_MS_GRAPH_RESOURCE_ = "MS_GRAPH_RESOURCE";
const CFG_MS_GRAPH_TENANT_ID = "MS_GRAPH_TENANT_ID";
const CFG_MS_GRAPH_GRANT_TYPE = "MS_GRAPH_GRANT_TYPE";

// Misc consts here
const CFG_TOKEN_GRACE_PERIOD = "TOKEN_GRACE_PERIOD";

const DEFAULT_TOKEN_GRACE_PERIOD = "5"; // In mins

process.on("unhandledRejection", error => {
  // Will print "unhandledRejection err is not defined"
  console.log("unhandledRejection", error);
});

// CNO365 class here
class CNO365 extends CNShell {
  // Properties here
  private _appId: string;
  private _clientSecret: string;
  private _resource: string;
  private _tenantId: string;
  private _grantType: string;

  private _token: string | undefined;
  private _tokenGracePeriod: number;
  private _tokenTimeout: NodeJS.Timeout;

  // Constructor here
  constructor(name: string, master?: CNShell) {
    super(name, master);

    this._appId = this.getRequiredCfg(CFG_MS_GRAPH_APP_ID, false, true);
    this._clientSecret = this.getRequiredCfg(
      CFG_MS_GRAPH_CLIENT_SECRET,
      false,
      true,
    );
    this._resource = this.getRequiredCfg(CFG_MS_GRAPH_RESOURCE_);
    this._tenantId = this.getRequiredCfg(CFG_MS_GRAPH_TENANT_ID, false, true);
    this._grantType = this.getRequiredCfg(CFG_MS_GRAPH_GRANT_TYPE);

    let gracePeriod = this.getCfg(
      CFG_TOKEN_GRACE_PERIOD,
      DEFAULT_TOKEN_GRACE_PERIOD,
    );

    this._tokenGracePeriod = parseInt(gracePeriod, 10) * 60 * 1000; // Convert to ms
  }

  // Abstract method implementations here
  async start(): Promise<boolean> {
    await this.renewToken();

    return true;
  }

  async stop(): Promise<void> {
    if (this._tokenTimeout !== undefined) {
      clearTimeout(this._tokenTimeout);
    }

    return;
  }

  async healthCheck(): Promise<boolean> {
    if (this._token === undefined) {
      return false;
    }

    return true;
  }

  // Private methods here
  private async renewToken(): Promise<void> {
    this.info("Renewing token now!");

    let data = {
      client_id: this._appId,
      client_secret: this._clientSecret,
      scope: `${this._resource}/.default`,
      grant_type: this._grantType,
    };

    let res = await this.httpReq({
      method: "post",
      url: `https://login.microsoftonline.com/${this._tenantId}/oauth2/v2.0/token`,
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
      },
      data: qs.stringify(data),
    }).catch(e => {
      this.error("Error while renewing token - (%s)", e);
    });

    if (res !== undefined) {
      this._token = res.data.access_token;

      let renewIn = res.data.expires_in * 1000 - this._tokenGracePeriod;

      this._tokenTimeout = setTimeout(() => this.renewToken(), renewIn);

      this.info(
        "Will renew token again in (%s) mins",
        Math.round(renewIn / 1000 / 60),
      );
    } else {
      // Try again in 1 minute
      this._token = undefined;
      this.info("Will try and renew token again in 1 min");
      this._tokenTimeout = setTimeout(() => this.renewToken(), 60 * 1000);
    }
  }
}

export { CNO365 };
