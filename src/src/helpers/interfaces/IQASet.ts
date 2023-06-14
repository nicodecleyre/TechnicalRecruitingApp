import { IQuestion } from "./IQuestion";

export interface IQASet {
    id: number,
    question: IQuestion;
    answer: string;
    score: number;
}