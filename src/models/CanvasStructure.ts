// https://learn.microsoft.com/en-us/graph/api/resources/canvaslayout?view=graph-rest-1.0
// https://graph.microsoft.com/v1.0/sites/2cf0e74f-9e9e-4108-87e5-eb283f7947a6/pages/59c9e772-3778-40ec-b8ce-75bbd742cb84/microsoft.graph.sitePage?$expand=canvasLayout
export interface CanvasStructure {
  "horizontalSections@odata.context": string
  horizontalSections: HorizontalSection[]
}

export interface HorizontalSection {
  layout: string
  id: string
  emphasis: string
  "columns@odata.context": string
  columns: Column[]
}

export interface Column {
  id: string
  width: number
  "webparts@odata.context": string
  webparts: Webpart[]
}

export interface Webpart {
  "@odata.type": string
  id: string
  innerHtml?: string
  webPartType?: string
  data?: Data
}

export interface Data {
  dataVersion: string
  description: string
  title: string
  properties: Properties
  serverProcessedContent: ServerProcessedContent
}

export interface Properties {
  imageSourceType?: number
  isAspectRatioLockedOnLoad?: boolean
  aspectRatioOnLoad?: number
  isOverlayTextVisible?: boolean
  imgHeight?: number
  imgWidth?: number
  alignment?: string
  fileName?: string
  cropX?: number
  cropY?: number
  cropWidth?: number
  cropHeight?: number
  fixAspectRatio?: boolean
  altText?: string
  advancedImageEditorData?: AdvancedImageEditorData
  overlayTextStyles?: OverlayTextStyles
  linkUrl?: string
  overlayText?: string
  siteId?: string
  webId?: string
  listId?: string
  uniqueId?: string
  resizeCoefficient?: number
  resizeDesiredWidth?: number
  cacheBuster?: string
  isOverlayTextEnabled?: boolean
  captionText?: string
  isMigrated?: boolean
  layoutId?: string
  shouldShowThumbnail?: boolean
  imageWidth?: number
  hideWebPartWhenEmpty?: boolean
  dataProviderId?: string
  iconPicker?: string
  "items@odata.type"?: string
  items?: Item[]
  listLayoutOptions?: ListLayoutOptions
  buttonLayoutOptions?: ButtonLayoutOptions
  waffleLayoutOptions?: WaffleLayoutOptions
  isDocumentLibrary?: boolean
  selectedListId?: string
  selectedListUrl?: string
  webRelativeListUrl?: string
  webpartHeightKey?: number
  selectedViewId?: string
  hideCommandBar?: boolean
  hideSeeAllButton?: boolean
}

export interface AdvancedImageEditorData {
  "@odata.type": string
  isAdvancedEdited: boolean
  originalSourceUrl?: string
  originalFileName?: string
  originalHeight?: number
  originalWidth?: number
}

export interface OverlayTextStyles {
  "@odata.type": string
  textColor: string
  isBold: boolean
  isItalic: boolean
  textBoxColor: string
  textBoxOpacity: number
  overlayColor: string
  overlayTransparency: number
  fontSize?: number
  position?: Position
}

export interface Position {
  "@odata.type": string
  offsetX: number
  offsetY: number
}

export interface Item {
  thumbnailType: number
  id: number
  description: string
  altText: string
  rawPreviewImageMinCanvasWidth: number
  shouldOpenInNewTab?: boolean
  fabricReactIcon?: FabricReactIcon
  sourceItem: SourceItem
}

export interface FabricReactIcon {
  "@odata.type": string
  iconName: string
}

export interface SourceItem {
  "@odata.type": string
  itemType: number
  fileExtension: string
  progId?: string
  guids?: Guids
}

export interface Guids {
  "@odata.type": string
  siteId: string
  webId: string
  listId: string
  uniqueId: string
}

export interface ListLayoutOptions {
  "@odata.type": string
  showDescription: boolean
  showIcon: boolean
}

export interface ButtonLayoutOptions {
  "@odata.type": string
  showDescription: boolean
  buttonTreatment: number
  iconPositionType: number
  textAlignmentVertical: number
  textAlignmentHorizontal: number
  linesOfText: number
}

export interface WaffleLayoutOptions {
  "@odata.type": string
  iconSize: number
  onlyShowThumbnail: boolean
}

export interface ServerProcessedContent {
  htmlStrings: any[]
  searchablePlainTexts: SearchablePlainText[]
  links: Link[]
  imageSources: ImageSource[]
}

export interface SearchablePlainText {
  key: string
  value: string
}

export interface Link {
  key: string
  value: string
}

export interface ImageSource {
  key: string
  value: string
}
