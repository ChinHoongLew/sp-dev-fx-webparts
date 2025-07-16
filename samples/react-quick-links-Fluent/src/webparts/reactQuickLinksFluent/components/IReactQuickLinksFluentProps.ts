import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IReactQuickLinksFluentProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;

  listName:string;
  groupBy:string;
  context:WebPartContext;

  quickLinkColor:string;
  quickLinkColor2:string;
  fontIconColor:string;
  margin:number;
  padding:number;
 maxWidth:number;
minHeight:number;
gridWidth:number;
}
