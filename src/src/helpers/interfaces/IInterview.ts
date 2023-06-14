import { ICandidate } from "./ICandidate";
import { IInterviewer } from "./IInterviewer";
import { IQASet } from "./IQASet";

export interface IInterview {
    id: number;
    dateOfInterview: Date;
    interviewer: IInterviewer;
    candidate: ICandidate;
    qaSet: IQASet[]
    overallScore: number,
    review: string
  }