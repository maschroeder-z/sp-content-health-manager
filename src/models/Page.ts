import { CanvasStructure } from "./CanvasStructure";
import { LinkInfo } from "./LinkInfo";

export interface Page {
  id: string;
  name?: string;
  title?: string;
  description?: string;
  webUrl?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  pageLayout: string;
  showComments: boolean;
  showRecommendedPages: boolean;
  thumbnailWebUrl?: string;
  canvasLayout?: CanvasStructure;
  // for analyze
  Links: LinkInfo[]|null;
  InProgress: boolean;
}


