/* eslint-disable @typescript-eslint/no-explicit-any */
import { IDomain } from "../../../helpers/interfaces/IDomain";
import { IInterview } from "../../../helpers/interfaces/IInterview";
import { IQuestion } from "../../../helpers/interfaces/IQuestion";
import { ShowInterviewScreen } from "../../../helpers/interfaces/ShowInterviewScreen";

export interface IInterviewsState {
    showScreen: ShowInterviewScreen,
    currentInterview: IInterview;
    interviews: IInterview[];
    dialogText: string;
    users: any[];
    domains: IDomain[];
    allQuestions: IQuestion[];
    loading: boolean;
    latestVersionSaved: boolean;
    openAIKey: string;
}
