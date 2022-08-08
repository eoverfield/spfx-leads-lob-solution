import { LeadView } from "..";
import { ILead } from "../../../../Lead";

export interface ILeadsState {
  loading: boolean;
  error: string | undefined;
  leads: ILead[];
  reminderCreating: boolean;
  reminderCreatingResult?: string;
  reminderDate?: Date;
  reminderDialogVisible: boolean;
  selectedLead?: ILead;
  submitCardDialogVisible: boolean;
  view: LeadView;
}