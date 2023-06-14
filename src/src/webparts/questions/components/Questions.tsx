/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from 'react';
import styles from './Questions.module.scss';
import { IQuestionsProps } from './IQuestionsProps';
import { SPFI } from '@pnp/sp';
import { IQuestionsState } from './IQuestionsState';
import { getSP } from '../../../helpers/pnpjsConfig';
import { Dialog, DialogFooter, DialogType, Dropdown, PrimaryButton, TextField } from 'office-ui-fabric-react';
import { IQuestion } from '../../../helpers/interfaces/IQuestion';
import { ShowQuestionScreen } from '../../../helpers/interfaces/showQuestionScreen';
import { Difficulty } from '../../../helpers/interfaces/difficulty';
import { Configuration, OpenAIApi } from 'openai';
import { IStatus } from '../../../helpers/interfaces/IStatus';

const emptyQuestion: IQuestion = {
    id: null,
    question: '',
    answer: '',
    domain: '',
    difficulty: null
}

export default class Questions extends React.Component<IQuestionsProps, IQuestionsState> {
  private _sp: SPFI;
  private static openai: OpenAIApi;

  constructor(props: IQuestionsProps){
    super(props);

    this.state = {
      showScreen: ShowQuestionScreen.list,
      questions: [],
      currentQuestionId: null,
      difficultyOptions: [],
      currentQuestion: Object.create(emptyQuestion),
      dialogText: null,
      openAIKey: null,
      status: null
    }

    this._sp = getSP(this.props.context);
  }

  public async componentDidMount(): Promise<void> {
    if(this._sp === null || this._sp === undefined){
      getSP(this.props.context);
    }

    await this.loadQuestions();
    this.loadDifficultyOptions();
    await this.loadConfigurationSettings();
  }

  private async loadConfigurationSettings(): Promise<void> {
    try{
      const items = await this._sp.web.lists.getByTitle('configuration').items.filter(`Title eq 'OpenAIKey'`)();
      if(items.length > 0 && items[0].Value !== ''){
        this.setState({
          openAIKey: items[0].Value
        });

        const key = new Configuration({
          apiKey: items[0].Value,
        });
        Questions.openai = new OpenAIApi(key);
      }
    } catch (error) {
      console.log(`Configuration setting not found. Error: ${error}`);
    }
  }

  private async loadQuestions(): Promise<void> {
    const questions: IQuestion[] = [];
    try{
      const items = await this._sp.web.lists.getByTitle('questions').items();

      for(const item of items){
        questions.push({
          id: item.Id,
          question: item.Title,
          difficulty: this.getDifficulty(item.Difficulty),
          answer: item.Answer,
          domain: item.Domain
        })
      }

      this.setState({
        questions: questions
      })
    } catch{
      this.setState({
        showScreen: ShowQuestionScreen.needConfig
      })
    }
  }

  private getDifficulty(difficulty: string): Difficulty{
    switch(difficulty){
      case 'easy':
        return Difficulty.easy
      case 'medium':
        return Difficulty.medium
      case 'hard':
        return Difficulty.hard
    }
  }

  private loadDifficultyOptions(): void {  
    const difficultyArray: { key: string, text: string }[] = Object.keys(Difficulty)
        .filter(key => isNaN(Number(key)))
        .map(key => ({ key, text: key }));
    
    this.setState({
      difficultyOptions: difficultyArray
    })
  }
  
  private onNewQuestionChange(field: string, event: any): void{
    let value: string = null;

    const question: IQuestion = this.state.currentQuestion;

    if(field !== 'Difficulty'){
      value = (event.target as HTMLInputElement).value;
    }


    switch(field){
      case 'Question':
        question.question = value
        break;
      case 'Answer':
        question.answer = value
        break;
      case 'Domain':
        question.domain = value
        break;
      case 'Difficulty':
        question.difficulty = this.getDifficulty(event);
    }

    this.setState({
      currentQuestion: question
    })
  }

  private async saveQuestion(): Promise<void> {
    let dialogText = '';

    this.setState({
      status: IStatus.loading
    });
    
    const spObject = {
      Title: this.state.currentQuestion.question,
      Answer: this.state.currentQuestion.answer,
      Domain: this.state.currentQuestion.domain,
      Difficulty: Difficulty[this.state.currentQuestion.difficulty]
    }

    if(this.state.showScreen === ShowQuestionScreen.new){
      await this._sp.web.lists.getByTitle('questions').items.add(spObject);
      dialogText = 'Question successfully saved';
    } else {
      await this._sp.web.lists.getByTitle('questions').items.getById(Number(this.state.currentQuestionId)).update(spObject);
      dialogText = 'Question successfully updated';
    }

    this.setState({
      dialogText: dialogText,
      currentQuestion: Object.create(emptyQuestion),
      currentQuestionId: null,
      status: IStatus.success
    })

    await this.loadQuestions();

    await this.sleep(5000);

    this.setState({
      showScreen: ShowQuestionScreen.list,
      dialogText: null
    })
  }

  private toggleHideDialog(): void {
    this.setState({
      dialogText: null,
      showScreen: ShowQuestionScreen.list
    })
  }

  private async removeQuestion(id: string): Promise<void> {
    await this._sp.web.lists.getByTitle('questions').items.getById(Number(id)).delete();

    this.setState({
      dialogText: 'Question succesfully removed',
    })

    await this.loadQuestions();

    await this.sleep(5000);

    this.setState({
      dialogText: null
    })
  }
  
  private async askOpenAi():Promise<void>{
    try{
      this.setState({
        status: IStatus.loading
      });

      const currentQuestion = this.state.currentQuestion;

      const response = await Questions.openai.createCompletion({
        model: "text-davinci-003",
        prompt: currentQuestion.question,
        max_tokens: 2000
      });

      const answer = response.data.choices[0].text.replace('\n\n','');

      currentQuestion.answer = answer;

      this.setState({
        currentQuestion: currentQuestion,
        status: IStatus.success
      })
    }catch(error: any){
      console.log(error);
      this.setState({
        dialogText: 'Something went wrong',
        status: null
      })
    }

  }

  private sleep(ms: number): Promise<unknown> {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  public render(): React.ReactElement<IQuestionsProps> {
    return (
      <section className={`${styles.questions} ${this.props.hasTeamsContext ? styles.teams : ''}`}>
        {this.state.showScreen === ShowQuestionScreen.needConfig && <section>
          👋 Hi There, We could not find a list named &apos;questions&apos;, please provide the necessary lists from the configuration tab
          </section>}
        <section className={styles.header}>{this.state.showScreen === ShowQuestionScreen.list && <PrimaryButton text='Create a new question' onClick={() => this.setState({
          showScreen: ShowQuestionScreen.new,
          currentQuestion: Object.create(emptyQuestion),
          currentQuestionId: null
        })}/>}
        {(this.state.showScreen !== ShowQuestionScreen.list && this.state.showScreen !== ShowQuestionScreen.needConfig) && <img alt='Close' src={require('../../../assets/close.png')} className={[styles.image,styles.close].join(' ')} onClick={() => this.setState({
                  currentQuestion: Object.create(emptyQuestion),
                  currentQuestionId: null,
                  showScreen: ShowQuestionScreen.list
                })}/>}
        </section>
        {this.state.showScreen === ShowQuestionScreen.list && <section>
          {this.state.questions.map(question => {
            return <section className={styles.question} key={question.id}>
              <section className={styles.questionInfo} onClick={() => this.setState({
                  currentQuestionId: question.id,
                  currentQuestion: this.state.questions.filter(x => x.id === question.id)[0],
                  showScreen: ShowQuestionScreen.view
                })}>
                <section className={styles.questionItem}>{question.question}</section>
                <section className={styles.questionItem}><b>{question.domain} ({Difficulty[question.difficulty]})</b></section>
              </section>
              <section className={styles.questionItem}>
                <img alt='Edit' src={require('../../../assets/edit.png')} className={styles.image} onClick={() => this.setState({
                  currentQuestionId: question.id,
                  currentQuestion: this.state.questions.filter(x => x.id === question.id)[0],
                  showScreen: ShowQuestionScreen.edit
                })}/>
              </section>
              <section className={styles.questionItem}>
              <img alt='Remove' src={require('../../../assets/remove.png')} className={styles.image} onClick={this.removeQuestion.bind(this, question.id)}/>
              </section>
            </section>
          })}
          </section>}
        
        {this.state.showScreen !== ShowQuestionScreen.list && <section className={this.state.showScreen === ShowQuestionScreen.view?styles.view:null}>
          <section className={styles.questionTitle}><TextField
              disabled={this.state.showScreen === ShowQuestionScreen.view}
              label='Question' value={this.state.currentQuestion.question}
              onChange={this.onNewQuestionChange.bind(this, 'Question')}/>
             {this.state.openAIKey && this.state.showScreen !== ShowQuestionScreen.view && <img alt='Remove' src={require('../../../assets/openai.svg')} className={styles.image} onClick={this.askOpenAi.bind(this)}/>}
          </section>
          <TextField
            disabled={this.state.showScreen === ShowQuestionScreen.view}
            label='Answer'
            value={this.state.currentQuestion.answer}
            multiline={true}
            onChange={this.onNewQuestionChange.bind(this, 'Answer')}/>
          <Dropdown
            disabled={this.state.showScreen === ShowQuestionScreen.view}
            label='Difficulty'
            options={this.state.difficultyOptions}
            selectedKey={this.state.currentQuestion.difficulty !== null?this.state.difficultyOptions.filter(x => x.key === Difficulty[this.state.currentQuestion.difficulty])[0].key:null}
            onChange={(event, selectedOption) => this.onNewQuestionChange('Difficulty', selectedOption.text)}/>
          <TextField
            disabled={this.state.showScreen === ShowQuestionScreen.view}
            label='Domain'
            value={this.state.currentQuestion.domain}
            onChange={this.onNewQuestionChange.bind(this, 'Domain')}/>
          {this.state.showScreen !== ShowQuestionScreen.view && <PrimaryButton
            className={styles.saveButton}
            text='Save'
            onClick={this.saveQuestion.bind(this)}/>}
            {this.state.status === IStatus.loading && <p><img alt='' src={require('../../../assets/loading.gif')} className={styles.statusImage}/></p>}
            {this.state.status === IStatus.success && <p><img alt='' src={require('../../../assets/checkmark.png')} className={styles.statusImage}/></p>}
          </section>}

          <Dialog
            hidden={this.state.dialogText === null}
            onDismiss={this.toggleHideDialog.bind(this)}
            dialogContentProps={{ type: DialogType.normal, title: 'Info', subText: this.state.dialogText}}
          >
            <DialogFooter>
              <PrimaryButton text='close' onClick={this.toggleHideDialog.bind(this)}/>
            </DialogFooter>
          </Dialog>
      </section>
    );
  }
}
