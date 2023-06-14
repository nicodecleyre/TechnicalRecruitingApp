import { Difficulty } from "./difficulty";

export interface IQuestion {
    id: number;
    question: string;
    answer: string;
    difficulty: Difficulty;
    domain: string;
}