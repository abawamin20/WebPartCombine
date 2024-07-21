import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IArticleDisplayProps {
  context: WebPartContext;
  groupId: string;
  setNames: string[];
  selectedView: string;
}
