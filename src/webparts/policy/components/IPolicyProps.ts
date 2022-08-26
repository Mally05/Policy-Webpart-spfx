import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IPolicyProps {
  description: string;
  context: WebPartContext;
  lists: string;
  fields: string[];
}
