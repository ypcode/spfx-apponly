import * as React from 'react';
import styles from './HardwareRequestForm.module.scss';
import { IHardwareRequestFormProps } from './IHardwareRequestFormProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { TextField, Dropdown, IDropdownOption, PrimaryButton } from "office-ui-fabric-react";
import { IHardwareRequest } from "../../../model/IHardwareRequest";
import { HardwareRequestService } from "../../../services/HardwareRequestService";



export default class HardwareRequestForm extends React.Component<IHardwareRequestFormProps, void> {

  private currentRequest: IHardwareRequest;
  private authenticated: boolean;

  constructor(props: IHardwareRequestFormProps) {
    super(props);
    this.currentRequest = {
      type: "",
      title: "",
      approved: false,
      quantity: 1,
      rejectionReason: "",
      remark: ""
    };
  }

  private submitRequest() {
    let service = HardwareRequestService.createForCurrentWeb();
    service.submitRequest(this.currentRequest);
  }

  public render(): React.ReactElement<IHardwareRequestFormProps> {
    return (
      <div className={styles.hardwareRequestForm}>
        <div className={styles.container}>
          <div className={`ms-Grid-row`}>
            <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <h2>Submit a new Hardware request</h2>
            </div>
          </div>
          <div className={`ms-Grid-row`}>
            <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <Dropdown label="Type" options={[
                { key: "Keyboard", text: "Keyboard" },
                { key: "Mouse", text: "Mouse" },
                { key: "Display", text: "Display" },
                { key: "Hard drive", text: "Hard drive" },
              ]} onChanged={v => this.currentRequest.type = v.key as string} />
            </div>
          </div>
          <div className={`ms-Grid-row`}>
            <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <TextField label="Quantity" onChanged={v => this.currentRequest.quantity = v} />
            </div>
          </div>
          <div className={`ms-Grid-row`}>
            <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <TextField multiline={true} rows={8} label="Remark" onChanged={v => this.currentRequest.remark = v} />
            </div>
          </div>
          <div className={`ms-Grid-row`}>
            <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <PrimaryButton text="Submit request" onClick={v => this.submitRequest()} />
            </div>
          </div>
          <div className={`ms-Grid-row`}>
            <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
            </div>
          </div>
        </div>
      </div>
    );
  }
}
