import { IStatus } from "../../../helpers/interfaces/IStatus";

export interface IConfigurationState {
  showProvisionButton: boolean;
  openAIKeyConfigItemID: string;
  openAIKey: string;
  provisionStatus: IStatus;
  dialogText: string;
}
