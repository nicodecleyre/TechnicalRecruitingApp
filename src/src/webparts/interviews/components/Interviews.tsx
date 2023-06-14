/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from 'react';
import styles from './Interviews.module.scss';
import { IInterviewsProps } from './IInterviewsProps';
import { IInterviewsState } from './IInterviewsState';
import { getSP } from '../../../helpers/pnpjsConfig';
import { SPFI } from '@pnp/sp';
import { ShowInterviewScreen } from '../../../helpers/interfaces/ShowInterviewScreen';
import { Checkbox, DatePicker, Dialog, DialogFooter, DialogType, Icon, PrimaryButton, TextField, initializeIcons } from 'office-ui-fabric-react';
import { IInterview } from '../../../helpers/interfaces/IInterview';
import { ICandidate } from '../../../helpers/interfaces/ICandidate';
import { IInterviewer } from '../../../helpers/interfaces/IInterviewer';
import "@pnp/sp/profiles";
import "@pnp/sp/site-users/web";
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { IDomain } from '../../../helpers/interfaces/IDomain';
import { Difficulty } from '../../../helpers/interfaces/difficulty';
import { IQuestion } from '../../../helpers/interfaces/IQuestion';
import { IQASet } from '../../../helpers/interfaces/IQASet';
import { IItemAddResult } from '@pnp/sp/items';
import { Configuration, OpenAIApi } from 'openai';

const emptyCandidate: ICandidate = {
  id: null,
  currentRole: '',
  email: '',
  name: '',
  shouldHire: false,
  yearsOfExperience: null
}

const emptyInterviewer: IInterviewer = {
  email: '',
  id: null,
  name: ''
}

const emptyInterview: IInterview = {
  id: null,
  candidate: {...emptyCandidate},
  dateOfInterview: null,
  interviewer: {...emptyInterviewer},
  overallScore: null,
  qaSet: [],
  review: ''
}

export default class Interviews extends React.Component<IInterviewsProps, IInterviewsState> {
  private _sp: SPFI;
  private static openai: OpenAIApi;
  
  constructor(props: IInterviewsProps){
    super(props);

    this.state = {
      showScreen: ShowInterviewScreen.list,
      currentInterview: JSON.parse(JSON.stringify(emptyInterview)),
      interviews: [],
      dialogText: null,
      users: [],
      domains: [],
      allQuestions: [],
      loading: false,
      latestVersionSaved: false,
      openAIKey: null
    }

    this._sp = getSP(this.props.context);
  }

  public async componentDidMount(): Promise<void> {
    initializeIcons();
    if(this._sp === null || this._sp === undefined){
      getSP(this.props.context);
    }

    await this.loadInterviews();
    await this.loadDomains();
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
        Interviews.openai = new OpenAIApi(key);
      }
    } catch (error) {
      console.log(`Configuration setting not found. Error: ${error}`);
    }
  }

  private async loadInterviews(): Promise<void> {
    const items = await this._sp.web.lists.getByTitle('interviews').items();
    const interviews: IInterview[] = [];
    for(const item of items){
      const candidate: ICandidate = await this.getCandidate(item.CandidateId);
      const interviewer = await this.getInterviewer(item.InterviewerId);
      interviews.push({
        id: item.Id,
        dateOfInterview: new Date(item.DateOfInterview),
        candidate: candidate,
        interviewer: interviewer,
        overallScore: item.Score,
        qaSet: [],
        review: item.Review,
      });
    }

    this.setState({
      interviews
    });
  }

  private async getCandidate(id: number): Promise<ICandidate> {
    const result = await this._sp.web.lists.getByTitle('candidates').items.getById(id)();
    const candidate:ICandidate = {
      id: result.Id,
      email: result.Title,
      name: result.Name,
      yearsOfExperience: result.YearsOfExperience,
      currentRole: result.CurrentRole,
      shouldHire: result.ShouldHire
    };

    return candidate;
  }

  private async getInterviewer(id: number): Promise<IInterviewer> {
    if(id){
      const interviewerResult = await this._sp.web.getUserById(id)();
      const interviewer:IInterviewer = {
        email: interviewerResult.Email,
        id: interviewerResult.Id,
        name: interviewerResult.Title
      };


      return interviewer;
    } else {
      return JSON.parse(JSON.stringify(emptyInterviewer));
    }
  }

  private async loadDomains(): Promise<void> {
    const questions = await this._sp.web.lists.getByTitle('questions').items.top(2000)();
    const uniqueDomains = new Set(questions.map((question) => question.Domain));
    const allQuestions: IQuestion[] = [];

    const domains: IDomain[] = [];
    uniqueDomains.forEach(domain => domains.push({
      name: domain,
      selected: false
    }));

    for(const question of questions){
      allQuestions.push({
        id: question.Id,
        question: question.Title,
        difficulty: this.getDifficulty(question.Difficulty),
        answer: question.Answer,
        domain: question.Domain
      })
    }

    this.setState({
      domains,
      allQuestions
    });
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

  private onInterviewChange(field: string, event: any): void{
    const value: string = (event.target as HTMLInputElement).value;
    const interview: IInterview = this.state.currentInterview;

    switch(field){
      case 'CandidateName':
        interview.candidate.name = value;
        break;
      case 'CandidateEmail':
        interview.candidate.email = value;
        break;
      case 'CandidateCurrentRole':
        interview.candidate.currentRole = value;
        break;
      case 'CandidateYearsOfExperience':
        interview.candidate.yearsOfExperience = Number(value);
        break;
    }

    this.setState({
      currentInterview: interview,
      latestVersionSaved: false
    })
  }

  private async _getPeoplePickerItems(items: any[]): Promise<void> {
    const username = items[0].secondaryText;
    const result = await this._sp.web.ensureUser(username);

    const interviewer: IInterviewer = {
      id: result.data.Id,
      email: username,
      name: items[0].text
    }
    const interview: IInterview = {...this.state.currentInterview, interviewer};
    interview.interviewer = interviewer;

    this.setState({
      currentInterview: interview,
      latestVersionSaved: false
    });
  }

  private _onDateChange(date: Date):void{
    const currentInterview: IInterview = this.state.currentInterview;
    currentInterview.dateOfInterview = date;
    this.setState({
      currentInterview,
      latestVersionSaved: false
    })
  }

  private onPreviousClick():void {
    let screen: ShowInterviewScreen
    switch(this.state.showScreen){
      case ShowInterviewScreen.closing:
        screen = ShowInterviewScreen.questions
        break;
      case ShowInterviewScreen.questions:
        screen = ShowInterviewScreen.domains
        break;
      case ShowInterviewScreen.domains:
        screen = ShowInterviewScreen.mainInfo
        break;
    }

    this.setState({
      showScreen: screen
    })
  }

  private onNextClick():void {
    let screen: ShowInterviewScreen;
    let dialogText: string = null;
    switch(this.state.showScreen){
      case ShowInterviewScreen.mainInfo:
        if(this.state.currentInterview.candidate.name === '' || this.state.currentInterview.candidate.email === ''){
          dialogText = 'Please fill in the required fields';
          screen = this.state.showScreen
        } else {
          screen = ShowInterviewScreen.domains
        }
        break;
      case ShowInterviewScreen.domains:
        screen = ShowInterviewScreen.questions
        break;
      case ShowInterviewScreen.questions:
        screen = ShowInterviewScreen.closing
        break;
    }

    this.setState({
      showScreen: screen,
      dialogText
    })
  }

  private onCheckboxClick(ev: any, checked: boolean): void{
    const domains = this.state.domains;
    const currentDomain = domains.filter(x => x.name === ev.target.ariaLabel)[0]
    currentDomain.selected = checked;
    this.setState({
      domains,
      latestVersionSaved: false
    });
  }

  private onShouldHireClick(ev: any, checked: boolean): void{
    const currentInterview = this.state.currentInterview;
    currentInterview.candidate.shouldHire = checked;
    this.setState({
      currentInterview,
      latestVersionSaved: false
    });
  }

  private async getRandomQuestions(): Promise<void> {
    const qaSet = [];
    const questions: IQuestion[] =[...this.state.allQuestions];
    const domains: string[] = this.state.domains.filter(x => x.selected).map(x => x.name);
    const experience: number = this.state.currentInterview.candidate.yearsOfExperience;

    const filteredQuestions = questions.filter((question) => domains.indexOf(question.domain) !== -1);
  
    const maxQuestions = Math.min(filteredQuestions.length, 5);
    const selectedQuestions: IQuestion[] = [];
  
    while (selectedQuestions.length < maxQuestions) {
      const randomIndex = Math.floor(Math.random() * filteredQuestions.length);
      const randomQuestion = filteredQuestions[randomIndex];
  
      // Adjust the difficulty selection based on years of experience
      if (experience <= 2 && randomQuestion.difficulty !== Difficulty.hard) {
        selectedQuestions.push(randomQuestion);
      } else if (experience >= 8 && randomQuestion.difficulty !== Difficulty.easy) {
        selectedQuestions.push(randomQuestion);
      } else if (experience > 2 && experience < 8) {
        selectedQuestions.push(randomQuestion);
      }

      questions.splice(randomIndex, 1);
    }

    for(const selectedQuestion of selectedQuestions){
      qaSet.push({
        id: null,
        question: selectedQuestion,
        answer: '',
        score: 0
      })
    }
  
    const currentInterview = this.state.currentInterview;

    currentInterview.qaSet = qaSet;

    this.setState({
      currentInterview,
      latestVersionSaved: false
    });

    await this.saveOrUpdateInterview();

  }

  private async saveOrUpdateInterview(): Promise<void> {
    this.setState({
      loading: true
    });

    const currentInterview = this.state.currentInterview;

    const candidateObject = {
      Title: this.state.currentInterview.candidate.email,
      Name: this.state.currentInterview.candidate.name,
      YearsOfExperience: this.state.currentInterview.candidate.yearsOfExperience,
      CurrentRole: this.state.currentInterview.candidate.currentRole,
      ShouldHire: this.state.currentInterview.candidate.shouldHire
    };

    if(currentInterview.candidate.id === null){
      // Add new candidate
      const candidate: IItemAddResult = await this._sp.web.lists.getByTitle('candidates').items.add(candidateObject);
      currentInterview.candidate.id = candidate.data.Id;

    } else {
      //update item
      await this._sp.web.lists.getByTitle('candidates').items.getById(Number(this.state.currentInterview.candidate.id)).update(candidateObject);
    }  

    const interviewObject = {
      DateOfInterview: this.state.currentInterview.dateOfInterview,
      InterviewerId: this.state.currentInterview.interviewer.id,
      CandidateId: currentInterview.candidate.id,
      Score: this.state.currentInterview.overallScore,
      Review: this.state.currentInterview.review
    }

    if(currentInterview.id  === null){
      // Add new item
      const interview:IItemAddResult = await this._sp.web.lists.getByTitle('interviews').items.add(interviewObject);
      currentInterview.id = interview.data.Id;
    } else {
      //update item
      await this._sp.web.lists.getByTitle('interviews').items.getById(Number(this.state.currentInterview.id)).update(interviewObject);
    }

    for(const question of this.state.currentInterview.qaSet){
      const interviewquestionObject = {
        InterviewId: currentInterview.id,
        QuestionId: question.question.id,
        Answer: question.answer,
        Score: question.score
      }

      if(question.id === null) {
        const q:IItemAddResult = await this._sp.web.lists.getByTitle('interviewquestionmapping').items.add(interviewquestionObject);
        question.id = q.data.Id;
        // Add new item
      } else {
        //update item
        await this._sp.web.lists.getByTitle('interviewquestionmapping').items.getById(Number(question.id)).update(interviewquestionObject);
      }
    }

    this.setState({
      loading: false,
      latestVersionSaved: true,
      currentInterview
    })
  }

  private ratingClicked(id: number, score: number): void{
    const currentInterview = this.state.currentInterview;

    currentInterview.qaSet.filter(x => x.id === id)[0].score = score;

    this.setState({
      currentInterview
    });

    this.calculateOverallRating();
  }

  private toggleHideDialog(): void {
    this.setState({
      dialogText: null
    })
  }

  private async rateIt(id: number):Promise<void>{
    try{
      const currentInterview = this.state.currentInterview;
      const currentQuestion = currentInterview.qaSet.filter(x => x.id === id)[0];

      const prompt = `if you had to rate how much this answer '${currentQuestion.answer}' corresponds to the actual answer '${currentQuestion.question.answer}', how much would you give if you gave a number between 0 and 10? Answer with the number only and judge harshly`;

      const response = await Interviews.openai.createCompletion({
        model: "text-davinci-003",
        prompt: prompt,
        max_tokens: 2000
      });

      const answer = response.data.choices[0].text.replace('\n\n','');

      let currentNumber = 0;
      const targetNumber = Number(answer);

      currentQuestion.score = 0;

      this.setState({
        currentInterview
      });

      while (currentNumber !== targetNumber) {        
        currentNumber++;

        currentQuestion.score = currentNumber;

        this.setState({
          currentInterview
        });
        
        await this.sleep(750);
      }

      this.calculateOverallRating();
    }catch(error: any){
      this.setState({
        dialogText: 'Something went wrong'
      })
    }
  }

  private updateAnswerCandidate(id: number, event: any): void{
    const value: string = (event.target as HTMLInputElement).value;

    const currentInterview = this.state.currentInterview;
    currentInterview.qaSet.filter(x => x.id === id)[0].answer = value;

    this.setState({
      currentInterview
    })
  }

  private updateFinalReview(event: any):void {
    const value: string = (event.target as HTMLInputElement).value;
    const currentInterview = this.state.currentInterview;

    currentInterview.review = value;

    this.setState({
      currentInterview
    });
  }

  private calculateOverallRating(): void {
    const currentInterview = this.state.currentInterview;

    if (currentInterview.qaSet.length !== 0) {
      const totalScore = currentInterview.qaSet.reduce((accumulator, item) => accumulator + item.score, 0);
      const averageScore = totalScore / currentInterview.qaSet.length;
      currentInterview.overallScore = averageScore;

      this.setState({
        currentInterview
      })
    }
  }

  private async showInterview(id: number): Promise<void> {
    const currentInterview = this.state.interviews.filter(x => x.id === id)[0];
    const qaSet: IQASet[] = [];

    const items = await this._sp.web.lists.getByTitle('interviewquestionmapping').items.filter(`InterviewId eq ${id}`).expand('Question').select('Id', 'Score', 'Answer', 'Question/ID')();
    for(const item of items){
      const q = await this._sp.web.lists.getByTitle('questions').items.getById(item.Question.ID)();
      qaSet.push({
        id: item.Id,
        answer: item.Answer,
        score: item.Score,
        question: {
          question: q.Title,
          answer: q.Answer,
          difficulty: this.getDifficulty(q.Difficulty),
          domain: q.Domain,
          id: q.Id
        }
      })
    }

    currentInterview.qaSet = qaSet;

    const activeDomains = new Set(qaSet.map((qaset) => qaset.question.domain));
    const domains = this.state.domains;
    for(const domain of domains){
      domain.selected = false;
    }
    activeDomains.forEach(domain => {
      domains.filter(x => x.name === domain)[0].selected = true
    });

    this.setState({
      currentInterview,
      domains,
      showScreen: ShowInterviewScreen.mainInfo
    });
  }

  private sleep(ms: number): Promise<unknown> {
    return new Promise(resolve => setTimeout(resolve, ms));
  }
  
  public render(): React.ReactElement<IInterviewsProps> {
    return (
      <section className={styles.interviews}>
        {this.state.showScreen === ShowInterviewScreen.needConfig && <section>
          👋 Hi There, We could not find a list named &apos;interviews&apos;, please provide the necessary lists from the configuration tab
          </section>}
        <section className={styles.header}>{this.state.showScreen === ShowInterviewScreen.list && <PrimaryButton text='Start a new interview' onClick={() => this.setState({
          showScreen: ShowInterviewScreen.mainInfo,
          currentInterview: JSON.parse(JSON.stringify(emptyInterview))
        })}/>}
        {(this.state.showScreen !== ShowInterviewScreen.list &&
          this.state.showScreen !== ShowInterviewScreen.needConfig &&
          this.state.showScreen !== ShowInterviewScreen.mainInfo) && <img alt='Previous' src={require('../../../assets/arrow.png')} className={[styles.image,styles.previous].join(' ')} onClick={this.onPreviousClick.bind(this)}/>}
        {(this.state.showScreen !== ShowInterviewScreen.list && this.state.showScreen !== ShowInterviewScreen.needConfig) && <img alt='Close' src={require('../../../assets/close.png')} className={[styles.image,styles.close].join(' ')} onClick={() => this.setState({
                  currentInterview: {...emptyInterview},
                  showScreen: ShowInterviewScreen.list
                })}/>}
        {(this.state.showScreen !== ShowInterviewScreen.list &&
          this.state.showScreen !== ShowInterviewScreen.needConfig &&
          this.state.showScreen !== ShowInterviewScreen.closing) && <img alt='Next' src={require('../../../assets/arrow.png')} className={[styles.image,styles.next].join(' ')} onClick={this.onNextClick.bind(this)}/>}
        </section>
        {this.state.showScreen === ShowInterviewScreen.list && <section className={styles.infoSection}>
          {this.state.interviews.map(interview => {
            return <section className={styles.interview} key={interview.id}>
              <section className={styles.interviewInfo} onClick={this.showInterview.bind(this, interview.id)}>
                <section className={styles.interviewListItem}>
                  <section>
                    <section className={styles.interviewItem}><b>{interview.candidate.name}</b></section>
                    <section className={styles.interviewItem}>{interview.candidate.currentRole}</section>
                  </section>
                  <section>
                    <section className={styles.interviewItem}><b>{`${interview.dateOfInterview.getDate()}/${interview.dateOfInterview.getMonth() + 1}/${interview.dateOfInterview.getFullYear()}`}</b></section>
                    <section className={styles.interviewItem}>{interview.interviewer.name}</section>
                  </section>
                </section>
              </section>
            </section>
          })}
          </section>}
        {this.state.showScreen === ShowInterviewScreen.mainInfo && <section className={styles.infoSection}>
          <TextField
            label='Name'
            value={this.state.currentInterview.candidate.name}
            onChange={this.onInterviewChange.bind(this, 'CandidateName')}
            required={true}/>
          <TextField
            label='E-mail'
            value={this.state.currentInterview.candidate.email}
            onChange={this.onInterviewChange.bind(this, 'CandidateEmail')}
            required={true}/>
          <TextField
            label='Current role'
            value={this.state.currentInterview.candidate.currentRole}
            onChange={this.onInterviewChange.bind(this, 'CandidateCurrentRole')}/>
          <TextField
            label='Years Of Experience'
            type='number'
            value={this.state.currentInterview.candidate.yearsOfExperience?this.state.currentInterview.candidate.yearsOfExperience.toString():''}
            onChange={this.onInterviewChange.bind(this, 'CandidateYearsOfExperience')}/>
          <PeoplePicker
            context={this.props.context}
            titleText="Interviewer"
            personSelectionLimit={1}
            showtooltip={true}
            principalTypes={[PrincipalType.User]}
            resolveDelay={1000}
            onChange={this._getPeoplePickerItems.bind(this)}
            defaultSelectedUsers={this.state.currentInterview.interviewer.email !== ''?[this.state.currentInterview.interviewer.email]:[]}/>
          <DatePicker
            label='Date of interview'
            onSelectDate={this._onDateChange.bind(this)}
            value={this.state.currentInterview.dateOfInterview}/>
          </section>}

          {this.state.showScreen === ShowInterviewScreen.domains && <section className={styles.infoSection}>
            {this.state.domains.map((domain) => {
              return <p key={domain.name}>
                  <Checkbox
                    checked={this.state.domains.length > 0 && this.state.domains.filter(x => x.name === domain.name)[0].selected}
                    label={domain.name}
                    disabled={this.state.currentInterview.qaSet.length > 0}
                    onChange={this.onCheckboxClick.bind(this)}/>
                </p>
            })}
            <PrimaryButton text='Create questions' onClick={this.getRandomQuestions.bind(this)} disabled={this.state.currentInterview.qaSet.length > 0}/>
            </section>}

          {this.state.showScreen === ShowInterviewScreen.questions && <section>{this.state.currentInterview.qaSet.length === 0 && <p>No questions found for this interview, please select the desired domains first and then generate the questions</p>}
            {this.state.currentInterview.qaSet.map((q: IQASet, i: number) => {
              return <section key={i} className={styles.questionItem}>
                <section className={styles.questionHeader}><b>{q.question.domain}</b> ({Difficulty[q.question.difficulty]})</section>
                <section><b>Question:</b></section>
                <section>{q.question.question}</section>
                <section className={styles.questionAnswer}><b>Answer:</b></section>
                <section>{q.question.answer}</section>
                <section><TextField
                            label='Answer candidate'
                            multiline={true}
                            onChange={this.updateAnswerCandidate.bind(this,q.id)}
                            value={q.answer}/></section>
                <section className={styles.questionFooter}>
                  {this.state.openAIKey && <PrimaryButton text='Rate it!' onClick={this.rateIt.bind(this, q.id)}/>}
                  <section>{[...Array(10)].map((_, index) => (
                    <Icon
                      className={styles.rating}
                      key={index}
                      iconName={(index <= q.score && q.score > 0) ? "FavoriteStarFill" : "FavoriteStar"}
                      onClick={this.ratingClicked.bind(this,q.id,index)}/>
                  ))}</section>
                </section>
              </section>
            })}
          </section>}
          {this.state.showScreen === ShowInterviewScreen.closing && <section className={styles.infoSection}>
            <TextField
              label='Final review'
              multiline={true}
              onChange={this.updateFinalReview.bind(this)}
              value={this.state.currentInterview.review}/>
            <Checkbox checked={this.state.currentInterview.candidate.shouldHire} label='Should Hire?' onChange={this.onShouldHireClick.bind(this)}/>
              <section>{[...Array(10)].map((_, index) => (
            <Icon
              className={styles.overallRating}
              key={index}
              iconName={(index <= this.state.currentInterview.overallScore && this.state.currentInterview.overallScore > 0) ? "FavoriteStarFill" : "FavoriteStar"}/>
            ))}</section>
            </section>}
          {this.state.showScreen !== ShowInterviewScreen.needConfig && this.state.showScreen !== ShowInterviewScreen.list && <section>
              <PrimaryButton text='Save' onClick={this.saveOrUpdateInterview.bind(this)} className={styles.saveButton}/>
              {this.state.loading && <p><img alt='' src={require('../../../assets/loading.gif')} className={styles.statusImage}/></p>}
              {this.state.latestVersionSaved && <p><img alt='' src={require('../../../assets/checkmark.png')} className={styles.statusImage}/></p>}
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
