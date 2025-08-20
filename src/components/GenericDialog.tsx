import { BaseDialog } from '@microsoft/sp-dialog';

export type DialogType = "success" | "warning";

export default class GenericDialog extends BaseDialog {
  private message: string;
  private type: DialogType;

  constructor(message: string, type: DialogType) {
    super();
    this.message = message;
    this.type = type;
  }

  public render(): void {
    const isSuccess = this.type === "success";

    const title = isSuccess ? "Success" : "Warning";
    const bgColor = isSuccess ? "#28a745" : "#f0ad4e";
    const icon = isSuccess ? "&#10004;" : "&#9888;";
    const iconColor = "white";

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
          background:${bgColor}; 
          display:flex; 
          align-items:center; 
          justify-content:center; 
          margin:0 auto 20px auto;
        ">
          <span style="color:${iconColor}; font-size:28px;">${icon}</span>
        </div>

        <h2 style="
          margin:0 0 10px 0; 
          font-size:20px; 
          font-weight:600; 
          color:#000;
        ">
          ${title}
        </h2>

        <p style="
          font-size:14px; 
          color:#555; 
          margin:0 0 20px 0;
          line-height:1.5;
          text-align:left;
        ">
          ${this.message}
        </p>

        <button type="button" 
                id="okButton" 
                style="
                  background:#fff;
                  border:1px solid rgb(138, 136, 134);
                  color:rgb(50, 49, 48);
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
      this.close().catch(() => { /* handle error */ });
    });

    this.domElement.querySelector('#closeButton')?.addEventListener('click', () => {
      this.close().catch(() => { /* handle error */ });
    });
  }
}
