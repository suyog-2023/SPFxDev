import * as React from 'react';
import { ChangeEvent } from 'react';
import styles from './HelloWorld.module.scss';
import { IHelloWorldProps } from './IHelloWorldProps';
// import { escape } from '@microsoft/sp-lodash-subset';
import { DatePicker, IDatePickerStrings } from 'office-ui-fabric-react/lib/DatePicker';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Checkbox } from '@fluentui/react';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Web } from "@pnp/sp/presets/all";
// import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';

export interface IStates {
  Items: any;
  ID: any;
  EmployeeName: any;
  EmployeeNameId: any;
  HireDate: any;
  JobDescription: any;
  Expertise: any;
  ExpertiseId: any;
  Payroll: any;
  HTML: any;
}

export default class CRUDReact extends React.Component<IHelloWorldProps, IStates> {
  constructor(props: IHelloWorldProps) {
    super(props);
    this.state = {
      Items: [],
      EmployeeName: "",
      EmployeeNameId: 0,
      ID: 0,
      HireDate: null,
      JobDescription: "",
      Expertise: "",
      ExpertiseId: 0,
      Payroll: "",
      HTML: []

    };
  }

  public async componentDidMount() {
    await this.getData();
  }

  public async getData() {

    let web = Web(this.props.webURL);
    const items: any[] = await web.lists.getByTitle("Employees").items.select("*", "EmployeeName/Title", "LookUP/Title").expand("EmployeeName", "LookUP").get();
    console.log(items);
    this.setState({ Items: items });
    let html = await this.getHTML(items);
    this.setState({ HTML: html });
  }

  public findData = (id: any): void => {
    var itemID = id;
    var allitems = this.state.Items;
    var allitemsLength = allitems.length;
    if (allitemsLength > 0) {
      for (var i = 0; i < allitemsLength; i++) {
        if (itemID == allitems[i].Id) {
          this.setState({
            ID: itemID,
            EmployeeName: allitems[i].EmployeeName.Title,
            EmployeeNameId: allitems[i].EmployeeNameId,
            HireDate: new Date(allitems[i].HireDate),
            JobDescription: allitems[i].JobDescription,
            Expertise: allitems[i].LookUP.Title,
            ExpertiseId: allitems[i].LookUPId,
            Payroll: allitems[i].Payroll
          });
        }
      }
    }

  }

  public async getHTML(items: any[]) {
    var tabledata =
      <table className={styles.tableStyles}>
        <thead>
          <tr>
            <th>Employee Name</th>
            <th>Hire Date</th>
            <th>Job Description</th>
            <th>Expertise</th>
            <th>Payroll</th>
          </tr>
        </thead>
        <tbody>
          {items && items.map((item, i) => {
            let payroll = item.Payroll;
            if (payroll == true) {
              payroll = "Yes";
            } else {
              payroll = "No";
            }
            return [
              <tr key={i} onClick={() => this.findData(item.ID)} className={styles.tr}>
                <td>{item.EmployeeName.Title}</td>
                <td>{FormatDate(item.HireDate)}</td>
                <td>{item.JobDescription}</td>
                <td>{item.LookUP.Title}</td>
                <td>{payroll}</td>
              </tr>
            ];
          })}
        </tbody>

      </table>;
    return await tabledata;
  }
  public _getPeoplePickerItems = async (items: any[]) => {

    if (items.length > 0) {

      this.setState({ EmployeeName: items[0].text });
      this.setState({ EmployeeNameId: items[0].id });
    }
    else {
      this.setState({ EmployeeNameId: "" });
      this.setState({ EmployeeName: "" });
    }
  }

  public _onChange(ev?: React.FormEvent<HTMLElement | HTMLInputElement>, isChecked?: boolean) {
    console.log(`The option has been changed to ${isChecked}.`);
  }

  private async SaveData() {
    let web = Web(this.props.webURL);
    await web.lists.getByTitle("Employees").items.add({
      EmployeeNameId: this.state.EmployeeNameId,
      HireDate: new Date(this.state.HireDate),
      JobDescription: this.state.JobDescription,
      LookUPId: this.state.ExpertiseId,
      Payroll: Boolean(this.state.Payroll)
    }).then(i => {
      console.log(i);
    });
    alert("Created Successfully");
    this.setState({ EmployeeName: "", HireDate: null, JobDescription: "", Expertise: "" });
    this.getData();
  }

  private async UpdateData() {
    let web = Web(this.props.webURL);
    await web.lists.getByTitle("Employees").items.getById(this.state.ID).update({

      EmployeeNameId: this.state.EmployeeNameId,
      HireDate: new Date(this.state.HireDate),
      JobDescription: this.state.JobDescription,
      LookUPId: this.state.ExpertiseId,
      Payroll: Boolean(this.state.Payroll)

    }).then(i => {
      console.log(i);
    });
    alert("Updated Successfully");
    this.setState({ EmployeeName: "", HireDate: null, JobDescription: "", Expertise: "" });
    this.getData();
  }

  private async DeleteData() {
    let web = Web(this.props.webURL);
    await web.lists.getByTitle("Employees").items.getById(this.state.ID).delete()
      .then(i => {
        console.log(i);
      });
    alert("Deleted Successfully");
    this.setState({ EmployeeName: "", HireDate: null, JobDescription: "" });
    this.getData();
  }

  public render(): React.ReactElement<IHelloWorldProps> {
    return (
      <div>
        <h1>CRUD Operations With ReactJs</h1>
        {this.state.HTML}
        <form className={styles.form}>
          <div>
            <Label>Employee Name</Label>
            <PeoplePicker
              context={this.props.context}
              personSelectionLimit={1}
              // defaultSelectedUsers={this.state.EmployeeName===""?[]:this.state.EmployeeName}
              required={false}
              onChange={this._getPeoplePickerItems}
              defaultSelectedUsers={[this.state.EmployeeName ? this.state.EmployeeName : ""]}
              showHiddenInUI={false}
              principalTypes={[PrincipalType.User]}
              resolveDelay={1000}
              ensureUser={true}
            />
          </div>
          <div>
            <Label>Hire Date</Label>
            <DatePicker maxDate={new Date()} allowTextInput={false} strings={DatePickerStrings} value={this.state.HireDate} onSelectDate={(e) => { this.setState({ HireDate: e }); }} ariaLabel="Select a date" formatDate={FormatDate} />
          </div>
          <div>
            <Label>Job Description</Label>
            <TextField
              value={this.state.JobDescription}
              multiline
              onChange={(event: ChangeEvent<HTMLInputElement>) => this.setState({ JobDescription: event.target.value })}
            />
          </div>
          <div>
            <Label>Experties</Label>
            <TextField
              value={this.state.Expertise}
              onChange={(event: ChangeEvent<HTMLInputElement>) => this.setState({ Expertise: event.target.value })}
            />
          </div>
          <div>
            <Label>Hired </Label> <Checkbox value={this.state.Payroll} onChange={(event: ChangeEvent<HTMLInputElement>) => this.setState({ Payroll: event.target.value })} />
          </div>
        </form>
        <div className={styles.buttonContainer}>
          <div className={styles.btns}><PrimaryButton text="Create" onClick={() => this.SaveData()} /></div>
          <div className={styles.btns}><PrimaryButton text="Update" onClick={() => this.UpdateData()} /></div>
          <div className={styles.btns}><PrimaryButton text="Delete" onClick={() => this.DeleteData()} /></div>
        </div>
      </div>
    );
  }
}
export const DatePickerStrings: IDatePickerStrings = {
  months: ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
  shortMonths: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
  days: ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
  shortDays: ['S', 'M', 'T', 'W', 'T', 'F', 'S'],
  goToToday: 'Go to today',
  prevMonthAriaLabel: 'Go to previous month',
  nextMonthAriaLabel: 'Go to next month',
  prevYearAriaLabel: 'Go to previous year',
  nextYearAriaLabel: 'Go to next year',
  invalidInputErrorMessage: 'Invalid date format.'
};
export const FormatDate = (date: any): string => {
  console.log(date);
  var date1 = new Date(date);
  var year = date1.getFullYear();
  var month = (1 + date1.getMonth()).toString();
  month = month.length > 1 ? month : '0' + month;
  var day = date1.getDate().toString();
  day = day.length > 1 ? day : '0' + day;
  return month + '/' + day + '/' + year;
};