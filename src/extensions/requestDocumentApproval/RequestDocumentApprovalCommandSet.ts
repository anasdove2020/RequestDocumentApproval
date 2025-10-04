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

export interface IRequestDocumentApprovalCommandSetProperties {
  sampleTextOne: string;
  sampleTextTwo: string;
}

export default class RequestDocumentApprovalCommandSet extends BaseListViewCommandSet<IRequestDocumentApprovalCommandSetProperties> {
  private _sharePointListService: ISharePointService;
  private _modalContainer: HTMLDivElement | null = null;

  public async onInit(): Promise<void> {
    this._sharePointListService = this.context.serviceScope.consume(SharePointService.serviceKey);

    const mainUrl = "https://mqxd.sharepoint.com";
    const mainSite = "Sandpit";
    const mainAprrovalList = "Request Approval";

    const mainSiteUrl = `${mainUrl}/sites/${mainSite}`;
    const mainRequestApprovalUrl = `/sites/${mainSite}/Lists/${mainAprrovalList}`;
    this._sharePointListService.setMainUrl(mainSiteUrl, mainRequestApprovalUrl);

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
      Dialog.alert(`Failed to submit approval request to SharePoint: ${errorMessage} Please check the SharePoint List exists and you have permissions to write to it.`).catch(() => { /* handle error */ });
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
}
