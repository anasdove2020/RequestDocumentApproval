import { BaseListViewCommandSet, type Command, type IListViewCommandSetExecuteEventParameters, type ListViewStateChangedEventArgs } from "@microsoft/sp-listview-extensibility";
import { Dialog } from "@microsoft/sp-dialog";
import * as React from "react";
import * as ReactDOM from "react-dom";
import { RequestApprovalModal } from "../../components";
import { ISelectedFile, IApprovalRequest } from "../../interfaces/IRequestApprovalModalProps";
import { ISharePointService } from "../../interfaces/ISharePointService";
import SharePointService from "../../services/SharePointService";
import GenericDialog from "../../components/GenericDialog";
import { CONTENT_TYPE } from "../../utils/constants";
import { ThemeProvider, ThemeChangedEventArgs, IReadonlyTheme } from '@microsoft/sp-component-base';

export interface IRequestDocumentApprovalCommandSetProperties {
  sampleTextOne: string;
  sampleTextTwo: string;
}

export default class RequestDocumentApprovalCommandSet extends BaseListViewCommandSet<IRequestDocumentApprovalCommandSetProperties> {
  private _sharePointListService: ISharePointService;
  private _modalContainer: HTMLDivElement | null = null;
  private _themeProvider: ThemeProvider;
  private _theme: IReadonlyTheme | undefined;

  public async onInit(): Promise<void> {
    this._sharePointListService = this.context.serviceScope.consume(SharePointService.serviceKey);

    const mainUrl = "https://mqxd.sharepoint.com";
    const mainSite = "Sandpit/test1";
    const mainAprrovalList = "Approval Request List";

    const mainSiteUrl = `${mainUrl}/sites/${mainSite}`;
    const mainRequestApprovalUrl = `/sites/${mainSite}/Lists/${mainAprrovalList}`;
    this._sharePointListService.setMainUrl(mainSiteUrl, mainRequestApprovalUrl);

    // Theme provider init
    this._themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);
    this._theme = this._themeProvider.tryGetTheme();
    this._applyThemeToIcon(this._theme);
    this._themeProvider.themeChangedEvent.add(this, this._handleThemeChanged);

    this._modalContainer = document.createElement("div");
    document.body.appendChild(this._modalContainer);
    const compareOneCommand: Command = this.tryGetCommand("REQUEST_APPROVAL_COMMAND");
    compareOneCommand.visible = false;
    this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);
    return Promise.resolve();
  }

  public onDispose(): void {
    if (this._modalContainer) {
      // eslint-disable-next-line @microsoft/spfx/pair-react-dom-render-unmount
      ReactDOM.unmountComponentAtNode(this._modalContainer);
      document.body.removeChild(this._modalContainer);
      this._modalContainer = null;
    }

    if (this._themeProvider) {
      this._themeProvider.themeChangedEvent.remove(this, this._handleThemeChanged);
    }

    super.onDispose();
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case "REQUEST_APPROVAL_COMMAND":
        this._handleRequestApproval(event);
        break;
      default:
        throw new Error("Unknown command");
    }
  }

  private _handleRequestApproval(_: IListViewCommandSetExecuteEventParameters): void {
    const selectedItems = this.context.listView.selectedRows;

    if (!selectedItems || selectedItems.length === 0) {
      Dialog.alert("Please select at least one document to request approval.").catch(() => { /* handle error */ });
      return;
    }

    if (selectedItems.length > 5) {
      const dialog = new GenericDialog("You have selected more than 5 items.\n\nPlease choose up to 5 items for request approval.", "warning");
      dialog.show().catch(() => { /* handle error */ });
      return;
    }

    const selectedFiles: ISelectedFile[] = selectedItems.map((item) => {
      const contentType = item.getValueByName("ContentType");
      const fsobjType = item.getValueByName("FSObjType");
      const isFolder = fsobjType === "1" || contentType === "Folder" || (contentType && contentType.indexOf("Folder") !== -1);

      return {
        name: item.getValueByName("FileLeafRef") || item.getValueByName("Title"),
        id: item.getValueByName("ID"),
        serverRelativeUrl: item.getValueByName("FileRef"),
        size: item.getValueByName("File_x0020_Size"),
        modified: item.getValueByName("Modified"),
        modifiedBy: item.getValueByName("Author")?.title || item.getValueByName("Author"),
        isFolder: isFolder,
        contentType: contentType
      };
    });

    const allowedContentTypes = [CONTENT_TYPE.AFCA_ACTIVITY_SET, CONTENT_TYPE.AFCA_DOC, CONTENT_TYPE.AFCA_PROCESS];
    const invalidFiles = selectedFiles.filter(
      f => !allowedContentTypes.includes(f.contentType)
    );

    if (invalidFiles.length > 0) {
      const invalidNames = invalidFiles.map(f => `- ${f.name}`).join("\n");
      const message = `The following files are not supported for request approval:\n${invalidNames}`;
      const htmlMessage = message.replace(/\n/g, "<br />");
      const dialog = new GenericDialog(htmlMessage, "warning");
      dialog.show().catch(() => { /* handle error */ });
      return;
    }
    
    this._showApprovalModal(selectedFiles);
  }

  private _showApprovalModal(selectedFiles: ISelectedFile[]): void {
    if (!this._modalContainer) return;

    const modalElement = React.createElement(RequestApprovalModal, {
      isOpen: true,
      selectedFiles: selectedFiles,
      context: this.context,
      onDismiss: () => {
        this._hideApprovalModal();
      },
      onSubmit: async (approvalRequest: IApprovalRequest) => {
        await this._processApprovalRequest(approvalRequest);
        this._hideApprovalModal();
      },
    });

    // eslint-disable-next-line @microsoft/spfx/pair-react-dom-render-unmount
    ReactDOM.render(modalElement, this._modalContainer);
  }

  private _hideApprovalModal(): void {
    if (!this._modalContainer) return;
    // eslint-disable-next-line @microsoft/spfx/pair-react-dom-render-unmount
    ReactDOM.unmountComponentAtNode(this._modalContainer);
  }

  private async _processApprovalRequest(approvalRequest: IApprovalRequest): Promise<void> {
    try {
      await this._sharePointListService.submitApprovalRequest(approvalRequest);
      const dialog = new GenericDialog(`Your approval request has been submitted successfully.`, "success");
      dialog.show()
            .then(() => { location.reload(); })
            .catch(() => { /* handle error */ });
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : String(error);

      if (error.message?.includes("locked")) {
        Dialog.alert(`Failed to submit approval request to SharePoint since the file is locked for editing.`).catch(() => { /* handle error */ });
      } else {
        Dialog.alert(`Failed to submit approval request to SharePoint: ${errorMessage} Please check the SharePoint List exists and you have permissions to write to it.`).catch(() => { /* handle error */ });
      }
      throw error; // Re-throw so modal can handle the error
    }
  }

  private _onListViewStateChanged = (_: ListViewStateChangedEventArgs): void => {
    const compareOneCommand: Command = this.tryGetCommand("REQUEST_APPROVAL_COMMAND");
    if (compareOneCommand) {
      const selectedCount = this.context.listView.selectedRows?.length || 0;
      compareOneCommand.visible = selectedCount >= 1;
    }
    this.raiseOnChange();
  };

  private _handleThemeChanged = (args: ThemeChangedEventArgs): void => {
    this._theme = args.theme;
    this._applyThemeToIcon(this._theme);
  };

  private _applyThemeToIcon(theme: IReadonlyTheme | undefined): void {
    const command = this.tryGetCommand("REQUEST_APPROVAL_COMMAND");
    if (!command) return;
  
    const color = theme?.palette?.themePrimary || "#0078d4"; // default SharePoint blue
  
    const svg = `
      <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"
          fill="none" stroke="${color}" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
        <!-- Document outline -->
        <path d="M9 2h6l5 5v13a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2z"/>
        <!-- Checkmark -->
        <polyline points="9 14 11 16 15 12"/>
      </svg>
    `;

    const base64 = btoa(svg);
    command.iconImageUrl = `data:image/svg+xml;base64,${base64}`;
  }
}
