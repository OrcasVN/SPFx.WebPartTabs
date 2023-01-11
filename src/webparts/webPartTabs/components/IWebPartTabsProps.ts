import { DisplayMode } from "@microsoft/sp-core-library";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export type tabStyle = 'links' | 'tabs';
export interface collectionTab {
  Title: string;
  WebPart: string;
  DisplayOrder: number;
  IconName: string;
  uniqueId: string;
}
export interface IWebPartTabsProps {
  wpContext: WebPartContext;
  tabStyle: tabStyle;
  collectionTabs: collectionTab[];
  displayMode: DisplayMode;
}
