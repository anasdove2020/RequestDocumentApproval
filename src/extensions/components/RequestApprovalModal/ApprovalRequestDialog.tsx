import { BaseDialog } from '@microsoft/sp-dialog';

export default class ApprovalRequestDialog extends BaseDialog {
  private message: string;

  constructor(message: string) {
    super();
    this.message = message;
  }

  public render(): void {
    this.domElement.innerHTML = `
      <div style="padding:20px;min-width:400px;max-width:900px;">
        <h2 style="margin-top:0; margin-bottom:10px;">Approval Request</h2>
        <p>${this.message}</p>
        <div style="display:flex; justify-content:flex-end;">
          <button type="button" 
                  class="ms-Button ms-Button--primary" 
                  id="okButton" 
                  style="min-width:120px;">
            <span class="ms-Button-label">OK</span>
          </button>
        </div>
      </div>
    `;

    this.domElement.querySelector('#okButton')?.addEventListener('click', () => {
      this.close().catch(() => {
        /* handle error */
      });
    });
  }
}
