export interface ILead {
  account: string;
  change: number;
  comments: ILeadComment[];
  createdBy: IPerson;
  createdOn: string;
  description?: string;
  id: string;
  percentComplete: number;
  requiresAttention?: boolean;
  title: string;
  url?: string;
}

export interface ILeadComment {
  comment: string;
  createdBy: IPerson;
  date: string;
}

export interface IPerson {
  email: string;
  name: string;
}