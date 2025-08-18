import { BaseDialog } from '@microsoft/sp-dialog';

export default class WarningDialog extends BaseDialog {
  private message: string;

  constructor(message: string) {
    super();
    this.message = message;
  }

  public render(): void {
    this.domElement.innerHTML = `
      <div style="
        background:#fff;
        padding:30px 20px 20px 20px;
        min-width:350px;
        max-width:420px;
        border-radius:12px;
        box-shadow:0 8px 20px rgba(0,0,0,0.15);
        text-align:center;
        font-family:Segoe UI, sans-serif;
        position:relative;
      ">
        <button id="closeButton" style="
          position:absolute;
          top:12px;
          right:12px;
          background:transparent;
          border:none;
          font-size:18px;
          font-weight:bold;
          color:#888;
          cursor:pointer;
        ">&times;</button>
        
        <div style="
          width:60px; 
          height:60px; 
          border-radius:50%; 
          background:#28a745; 
          display:flex; 
          align-items:center; 
          justify-content:center; 
          margin:0 auto 20px auto;
        ">
          <span style="color:white; font-size:28px;">&#10004;</span>
        </div>

        <h2 style="
          margin:0 0 10px 0; 
          font-size:20px; 
          font-weight:600; 
          color:#000;
        ">
          Success
        </h2>

        <p style="
          font-size:14px; 
          color:#555; 
          margin:0 0 20px 0;
          line-height:1.5;
        ">
          ${this.message}
        </p>

        <button type="button" 
                id="okButton" 
                style="
                  background:#28a745;
                  border:none;
                  color:white;
                  padding:12px 0;
                  font-size:14px;
                  font-weight:600;
                  width:100%;
                  border-radius:6px;
                  cursor:pointer;
                ">
          OK
        </button>
      </div>
    `;

    this.domElement.querySelector('#okButton')?.addEventListener('click', () => {
      this.close().catch(() => {
        /* handle error */
      });
    });

    this.domElement.querySelector('#closeButton')?.addEventListener('click', () => {
      this.close().catch(() => {
        /* handle error */
      });
    });
  }
}
