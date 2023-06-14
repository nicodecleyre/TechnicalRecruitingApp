/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from 'react';
import styles from './Configuration.module.scss';
import { IConfigurationProps } from './IConfigurationProps';
import { Dialog, DialogFooter, DialogType, PrimaryButton, TextField } from 'office-ui-fabric-react';
import { getSP } from '../../../helpers/pnpjsConfig';
import { IConfigurationState } from './IConfigurationState';
import { SPFI } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/views";
import { DateTimeFieldFormatType, FieldUserSelectionMode } from "@pnp/sp/fields/types";
import { IStatus } from '../../../helpers/interfaces/IStatus';

export default class Configuration extends React.Component<IConfigurationProps, IConfigurationState> {
  private _sp: SPFI;
  private static readonly lists: string[] = ['questions','candidates','interviews','configuration'];
  private static readonly group: string = 'TechnicalReqruitingApp';

  constructor(props: IConfigurationProps){
    super(props);

    this.state = {
      showProvisionButton: false,
      openAIKeyConfigItemID: '',
      openAIKey: '',
      provisionStatus: null,
      dialogText: null
    }

    this._sp = getSP(this.props.context);
  }

  private openAIChange(event: any): void{
    const value = (event.target as HTMLInputElement).value;

    this.setState({
      openAIKey: value
    })
  }

  public async componentDidMount(): Promise<void> {
    if(this._sp === null || this._sp === undefined){
      getSP(this.props.context);
    }

    await this.loadConfigurationSettings();
    const configSucceed: boolean = await this.checkListsExistence();
    if(!configSucceed){
      this.setState({
        showProvisionButton: true
      })
    }
  }

  private async loadConfigurationSettings(): Promise<void> {
    try{
      const items = await this._sp.web.lists.getByTitle('configuration').items.filter(`Title eq 'OpenAIKey'`)();
      if(items.length > 0){
        this.setState({
          openAIKey: items[0].Value,
          openAIKeyConfigItemID: items[0].Id
        });
      }
    } catch (error) {
      console.log(`Configuration setting not found. Error: ${error}`);
    }
  }

  private async saveOpenAIKey(): Promise<void> {
    const isValidKey = this.isValidOpenAIKey(this.state.openAIKey);

    if(!isValidKey){
      this.setState({
        dialogText: 'the key is you provided is not a valid open AI key. Expected key format: starts with "sk-" followed by alphanumeric characters'
      })
    }

    if(!this.state.openAIKeyConfigItemID){
      await this._sp.web.lists.getByTitle('configuration').items.add({
        Title: 'OpenAIKey',
        Value: this.state.openAIKey
      });
    } else {
      await this._sp.web.lists.getByTitle('configuration').items.getById(Number(this.state.openAIKeyConfigItemID)).update({
        Title: 'OpenAIKey',
        Value: this.state.openAIKey
      });
    }

    this.setState({
      dialogText: 'Key successfully saved!'
    });

    await this.sleep(5000);

    this.setState({
      dialogText: null
    });
  }

  private isValidOpenAIKey(key: string): boolean {
    // Expected key format: starts with "sk-" followed by alphanumeric characters
    const keyRegex = /^sk-[A-Za-z0-9]+$/;
  
    return keyRegex.test(key);
  }

  private async checkListsExistence():Promise<boolean> {
      for (const listTitle of Configuration.lists) {  
        try {
          await this._sp.web.lists.getByTitle(listTitle)();
        } catch (error) {
          console.log(`List '${listTitle}' does not exist.`);
          return false;
        }
      }
  
      console.log('All lists exist.');
      return true;
  }

  private async createLists(): Promise<void> {
    try {
      this.setState({
        provisionStatus: IStatus.loading
      })

      // Create 'questions' list
      const questionsList = await this._sp.web.lists.add('questions', 'questions', 100);
      await this._sp.web.lists.getByTitle('questions').fields.getByTitle('Title').update({ Title: 'Question'});
      const questionsAnswer = await this._sp.web.lists.getByTitle('questions').fields.addMultilineText('Answer', {RichText: false, NumberOfLines: 6, Group: Configuration.group});
      const questionsDifficulty = await this._sp.web.lists.getByTitle('questions').fields.addChoice('Difficulty', {Choices: ['easy', 'medium', 'hard'], FillInChoice: false, Group: Configuration.group});
      const questionsDomain = await this._sp.web.lists.getByTitle('questions').fields.addChoice('Domain', {Choices: [], FillInChoice: true, Group: Configuration.group});

      // Create 'candidates' list
      const candidatesList = await this._sp.web.lists.add('candidates', 'candidates', 100);
      await this._sp.web.lists.getByTitle('candidates').fields.getByTitle('Title').update({ Title: 'E-mail'});
      const candidatesName = await this._sp.web.lists.getByTitle('candidates').fields.addText('Name', {Group: Configuration.group});
      const candidatesYearsOfExperience = await this._sp.web.lists.getByTitle('candidates').fields.addNumber('YearsOfExperience', {MaximumValue: 100, Group: Configuration.group});
      const candidatesCurrentRole = await this._sp.web.lists.getByTitle('candidates').fields.addText('CurrentRole', {Group: Configuration.group});
      const candidatesShouldHire = await this._sp.web.lists.getByTitle('candidates').fields.addBoolean('ShouldHire', {Group: Configuration.group});

      // Create 'interviews' list
      const interviewList = await this._sp.web.lists.add('interviews', 'interviews', 100);
      await this._sp.web.lists.getByTitle('interviews').fields.getByTitle('Title').update({Required: false});
      const interviewDateOfInterview = await this._sp.web.lists.getByTitle('interviews').fields.addDateTime('DateOfInterview', {DisplayFormat: DateTimeFieldFormatType.DateTime, Group: Configuration.group})
      const interviewInterviewer = await this._sp.web.lists.getByTitle('interviews').fields.addUser('Interviewer', { SelectionMode: FieldUserSelectionMode.PeopleOnly, Group: Configuration.group})
      const interviewCandidate = await this._sp.web.lists.getByTitle('interviews').fields.addLookup('Candidate', {LookupListId: candidatesList.data.Id});
      const interviewScore = await this._sp.web.lists.getByTitle('interviews').fields.addNumber('Score', {MaximumValue: 10, Group: Configuration.group});
      const interviewReview = await this._sp.web.lists.getByTitle('interviews').fields.addMultilineText('Review', {RichText: false, NumberOfLines: 6, Group: Configuration.group});


      // Create 'interviews/questions' mapping list
      await this._sp.web.lists.add('interviewquestionmapping', 'interviewquestionmapping', 100);
      await this._sp.web.lists.getByTitle('interviews').fields.getByTitle('interviewquestionmapping').update({Required: false});
      const interviewquestionInterview = await this._sp.web.lists.getByTitle('interviewquestionmapping').fields.addLookup('Interview', {LookupListId: interviewList.data.Id});
      const interviewquestionQuestion = await this._sp.web.lists.getByTitle('interviewquestionmapping').fields.addLookup('Question', {LookupListId: questionsList.data.Id});
      const interviewquestionAnswer = await this._sp.web.lists.getByTitle('interviewquestionmapping').fields.addMultilineText('Answer', {NumberOfLines: 6, RichText: false, Group: Configuration.group});
      const interviewquestionScore = await this._sp.web.lists.getByTitle('interviewquestionmapping').fields.addNumber('Score', {MaximumValue: 10, Group: Configuration.group});

      // Create 'configuration' list
      await this._sp.web.lists.add('configuration', 'configuration', 100);
      await this._sp.web.lists.getByTitle('configuration').fields.getByTitle('Title').update({ Title: 'key'});
      await this._sp.web.lists.getByTitle('configuration').fields.addText('Value', {Group: Configuration.group});

      // Add fields to defaultview
      await this._sp.web.lists.getByTitle('questions').defaultView.fields.add(questionsAnswer.data.InternalName);
      await this._sp.web.lists.getByTitle('questions').defaultView.fields.add(questionsDifficulty.data.InternalName);
      await this._sp.web.lists.getByTitle('questions').defaultView.fields.add(questionsDomain.data.InternalName);
      await this._sp.web.lists.getByTitle('candidates').defaultView.fields.add(candidatesName.data.InternalName);
      await this._sp.web.lists.getByTitle('candidates').defaultView.fields.add(candidatesYearsOfExperience.data.InternalName);
      await this._sp.web.lists.getByTitle('candidates').defaultView.fields.add(candidatesCurrentRole.data.InternalName);
      await this._sp.web.lists.getByTitle('candidates').defaultView.fields.add(candidatesShouldHire.data.InternalName);
      await this._sp.web.lists.getByTitle('interviews').defaultView.fields.add(interviewDateOfInterview.data.InternalName);
      await this._sp.web.lists.getByTitle('interviews').defaultView.fields.add(interviewInterviewer.data.InternalName);
      await this._sp.web.lists.getByTitle('interviews').defaultView.fields.add(interviewCandidate.data.InternalName);
      await this._sp.web.lists.getByTitle('interviews').defaultView.fields.add(interviewScore.data.InternalName);
      await this._sp.web.lists.getByTitle('interviews').defaultView.fields.add(interviewReview.data.InternalName);
      await this._sp.web.lists.getByTitle('interviewquestionmapping').defaultView.fields.add(interviewquestionInterview.data.InternalName);
      await this._sp.web.lists.getByTitle('interviewquestionmapping').defaultView.fields.add(interviewquestionQuestion.data.InternalName);
      await this._sp.web.lists.getByTitle('interviewquestionmapping').defaultView.fields.add(interviewquestionAnswer.data.InternalName);
      await this._sp.web.lists.getByTitle('interviewquestionmapping').defaultView.fields.add(interviewquestionScore.data.InternalName);
      console.log('Lists and fields created successfully.');

      this.setState({
        provisionStatus: IStatus.success,
        dialogText: 'List and fields succesfully provided',
      })

      await this.sleep(5000);

      this.setState({
        showProvisionButton: false,
        dialogText: null
      })

    } catch (error) {
      this.setState({
        provisionStatus: null,
        dialogText: `An error occurred: ${error.toString()}`
      })
      console.log('An error occurred:', error);
    }
  }

  private toggleHideDialog(): void {
    this.setState({
      dialogText: null
    })
  }
  
  private sleep(ms: number): Promise<unknown> {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  public render(): React.ReactElement<IConfigurationProps> {
    return (
      <section className={`${styles.configuration} ${this.props.hasTeamsContext ? styles.teams : ''}`}>
        {!this.state.showProvisionButton && <section>
            <TextField label='Open AI Key' type='password' value={this.state.openAIKey} onChange={this.openAIChange.bind(this)}/>
            <PrimaryButton className={styles.button} text='save' onClick={this.saveOpenAIKey.bind(this)}/>
          </section>}
        {this.state.showProvisionButton && <section>
          <p>To use this app you need some necessary lists. You can provide this by pressing the button below</p>
          <p><PrimaryButton text='Provision lists' onClick={this.createLists.bind(this)} className={styles.button}/></p>
          {this.state.provisionStatus === IStatus.loading && <p><img alt='' src={require('../../../assets/loading.gif')} className={styles.loadingImage}/></p>}
          {this.state.provisionStatus === IStatus.success && <p><img alt='' src={require('../../../assets/checkmark.png')} className={styles.successImage}/> You will be redirected in a few seconds</p>}
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
