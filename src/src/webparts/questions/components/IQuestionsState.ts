import { IDropdownOption } from "office-ui-fabric-react";
import { IQuestion } from "../../../helpers/interfaces/IQuestion";
import { ShowQuestionScreen } from "../../../helpers/interfaces/showQuestionScreen";
import { IStatus } from "../../../helpers/interfaces/IStatus";

export interface IQuestionsState {
  showScreen: ShowQuestionScreen;
  questions: IQuestion[];
  currentQuestionId: number;
  difficultyOptions: IDropdownOption[];
  currentQuestion: IQuestion;
  dialogText: string;
  openAIKey: string;
  status: IStatus;
}
