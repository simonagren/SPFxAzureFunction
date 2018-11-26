import { AadHttpClient } from "@microsoft/sp-http";

export interface IHelloWorldProps {
  description: string;
  client: AadHttpClient;
}
