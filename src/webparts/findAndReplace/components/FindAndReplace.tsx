import * as React from 'react';
import styles from './FindAndReplace.module.scss';
import { IFindAndReplaceProps } from './IFindAndReplaceProps';
import { IFindAndReplaceStates } from './IFindAndReplaceStates';
import { escape, times } from '@microsoft/sp-lodash-subset';
import { FindAndReplaceContext } from "../FindAndReplaceWebPart";
import { sp , Web  } from "@pnp/sp/presets/all";
import { DefaultButton, PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Dialog, DialogFooter, DialogType } from 'office-ui-fabric-react/lib/Dialog';
import { Dropdown } from 'office-ui-fabric-react';
import "@pnp/sp/fields";
export default class FindAndReplace extends React.Component<IFindAndReplaceProps, IFindAndReplaceStates> {

  private ThisPageContext: any;
  private FandRContext: any = { ...FindAndReplaceContext };
  private web: any = null;
  private currentUserMail: string = "";
  private currentUserName: string = "";
  private contextfrompage: any;
  private StoreData: any = [];
  private ColumnData: any = [];
  private FindText: any = [];
  private URL: any = null;
  private ListData: any = [];
  private selectedList: any = [];
  private SelectedColumn: any = [];
  private SelectedFindText: any = "";
  private ReplaceWithText: any = "";
  constructor(props) {
    super(props);
    sp.setup({
      spfxContext: this.props.context
    });

    this.contextfrompage = this.FandRContext;
    this.ThisPageContext = this.props.context;
    this.web = Web(this.props.SiteName);
    this.URL = this.ThisPageContext._pageContext.site.absoluteUrl
    this.currentUserMail = this.ThisPageContext._pageContext.user.displayName;
    this.currentUserName = this.ThisPageContext._pageContext.user.email;

    this.state = {
      ReplaceText: "",
      FindText: "",
      ColumnName: "",
      ListName: "",
      ErrorMessagelist: "",
      hDialog: true,
      DHeader: "",
      DMessage: "",
      HideFooter: false,
      Listoptions: [],
      Columnoptions: [],
      FindTextOptions: []
    };
    this.buttonclick = this.buttonclick.bind(this);
    this.Replaceclick = this.Replaceclick.bind(this);
  }



  public async componentDidMount() {

    await this.LoadListData();
  }

  private buttonclick() {

    this.setState({
      DHeader: "Warning ",
      DMessage: "Do you want to replace items ?", hDialog: false
    })
  }

  private async LoadListData() {
    this.ListData = await this.web.lists.filter("(BaseTemplate ne 101) and (Hidden eq false)").select("Title").get().then((lists) => {
      var ListNames = [];
      lists.forEach(element => {
        if (element.Title != "Site Pages" && element.Title != "Site Assets") {
          ListNames.push({ key: element.Title, text: element.Title });
        }
      });
      this.setState({ Listoptions: ListNames });
      console.log("Load List Data : ",this.ListData);
    });
  }

  private LoadColumnData() {

    this.web.lists.getByTitle(this.selectedList).fields.get().then(async (AllListData) => {
      var ListColumnName = [];

      AllListData.forEach(element => {
        if (element.TypeDisplayName == "Single line of text" &&
          element.Title != "Copy Source" && element.Title != "Version" && element.Title != "File Type" &&
          element.Title != "Compliance Asset Id") {
          ListColumnName.push({ key: element.InternalName, text: element.InternalName });
        }
      });
      this.setState({ Columnoptions: ListColumnName.sort((a, b) => a.text.localeCompare(b.text)) });
      console.log("Column data : ", ListColumnName);

      // important to include all columns
      // sp.web.lists.getByTitle(this.selectedList).items
      //   .select("*").getAll().then(async (AllListData) => {
      //     var ListColumnName = [];
      //     var a = Object.getOwnPropertyNames(AllListData[0]); // all columns from list
      //     for (var i = 0; i < a.length; i++) {
      //       ListColumnName.push({ key: a[i], text: a[i] });
      //     }
      //          //this.StoreData = AllListData;

      //     this.setState({ Columnoptions: ListColumnName.sort((a, b) => a.text.localeCompare(b.text)) });
      //     console.log("Column data : ", await sp.web.lists.getByTitle(this.selectedList).fields.get());
      //     //this.columncall();
      //   });

    });
  }

  private LoadFindTextOptions() {
    var FindTextArray = [];
    this.web.lists.getByTitle(this.selectedList).items
      .select(this.SelectedColumn,"Id").getAll().then((AllColumnData) => {
        var flags = [], l = AllColumnData.length, i: number;
        var columnnametempstore = this.SelectedColumn;
        for (i = 0; i < l; i++) {
          if (flags[AllColumnData[i][columnnametempstore]]) continue;
          flags[AllColumnData[i][columnnametempstore]] = true;
          FindTextArray.push({ 
            key: AllColumnData[i][columnnametempstore], 
            text: AllColumnData[i][columnnametempstore]});
        }    
      });
     
    this.setState({ FindTextOptions: FindTextArray.sort((a, b) => a.text.localeCompare(b.text)) });
    console.log("Find Text Options : ", FindTextArray);
    
  }

  private getAllTextID()
  {
  
    this.web.lists.getByTitle(this.selectedList).items
      .select(this.SelectedColumn,"*").getAll().then((AllData) => {
        AllData.forEach(element => {

        if(element[this.SelectedColumn] == this.SelectedFindText)
        {
          this.FindText.push({ "ID": element.ID , "Title": element[this.SelectedColumn] });
        }
      });
        console.log("ID Data : ", this.FindText);
        }    
      );
  
  }
  private Replaceclick() {
    try {

      this.setState({ DHeader: "Please wait...", DMessage: "Your request is in process", hDialog: false, HideFooter: true });
      this.Replacedata();
      this.selectedList = null;
      this.SelectedColumn = null;
      this.ReplaceWithText = null;
      this.setState({FindTextOptions : [], Columnoptions : [],Listoptions : [], ReplaceText : null});
      
    }
    catch (ex) { console.log(ex) };
  }
  private async Replacedata() {

    try {
      var listname = this.selectedList;
      var InternalnameofColumn = this.SelectedColumn;
      var body = { [InternalnameofColumn]: this.ReplaceWithText };
      var storeupdateddata = [];
      this.setState({ DHeader: "Please wait...", DMessage: `Replacing ${this.FindText.length} items with your Find Text title "${this.state.FindText}"`, hDialog: false, HideFooter: true });

      let batch = this.web.createBatch();

      for (var i = 0; this.FindText[i]; i++) {
        // abc.push(this.web.lists.getByTitle(this.state.ListName).items.inBatch(batch).getById(this.FindText[i].ID).update(body))      

        this.web.lists.getByTitle(listname).items.inBatch(batch).getById(this.FindText[i].ID).update(body).then(() => {
          storeupdateddata.push(this.FindText[i]);
          this.setState({ DMessage: `Replacing ${this.FindText.length}/${storeupdateddata.length} items`, HideFooter: true, hDialog: false });
          if (this.FindText.length === storeupdateddata.length) {
            this.setState({
              DHeader: "Success", DMessage: `Replaced ${storeupdateddata.length}/${this.FindText.length} items. You can close the window`, hDialog: false, HideFooter: true
              , ListName: "", ColumnName: "", FindText: "", ReplaceText: ""
            });
           
          }
        });
      }
      await batch.execute();
      
    }
    catch (ex) {
      this.setState({ DHeader: "Error", DMessage: ` ${ex.message} , You can try again.`, hDialog: false, HideFooter: true });
    }

    // this.setState({ DHeader: "Success",DMessage : "Replaced"});
  }

  // private columncall() {
  //   for (var i = 0; this.StoreData[i]; i++) {
  //     if (this.StoreData[i].hasOwnProperty(this.state.ColumnName)) {
  //       var columnName = this.state.ColumnName;
  //       //this.ColumnData.push({ "ID" : this.StoreData[i].ID , "ColumnData": this.StoreData[i][columnName]});
  //       this.ColumnData.push({ "ID": this.StoreData[i].ID, "ColumnData": this.StoreData[i][columnName] });
  //     }
  //   }
  //   console.log("Column Data : ", this.ColumnData);
  //   this.findText();
  // }

  // private findText() {

  //   for (var i = 0; this.ColumnData[i]; i++) {
  //     if (this.ColumnData[i]["ColumnData"].includes(this.state.FindText)) {
  //       this.FindText.push({ "ID": this.ColumnData[i].ID, "MatchedText": this.ColumnData[i]["ColumnData"] });
  //     }
  //   }
  //   // this.setState({ DHeader: "Please wait...", DMessage: `We found ${this.FindText.length} items which match with your Find Text title "${this.state.FindText}"`, hDialog: false, HideFooter: true });

  //   console.log("Find Text matched Data : ", this.FindText);


  //   this.Replacedata();
  // }

  


  public render(): React.ReactElement<IFindAndReplaceProps> {
    return (
      <div className={styles.findAndReplace}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <Dropdown
              selectedKey={this.selectedList}
                options={this.state.Listoptions}
                label="List Name"
                placeholder="Select List..."
                onChange={(ev, item: any) => {
                  this.selectedList = item.key;
                  this.setState({ ListName: item.key });
                  this.LoadColumnData();
                }}
              >
              </Dropdown>


              {/* 
              <TextField label="List Name"
                onChange={(ev, text: any) => {
                  this.setState({ ListName: text });
                }} /> */}
            </div>
            <div className={styles.column}>

              <Dropdown
              selectedKey={this.SelectedColumn}
                options={this.state.Columnoptions}
                label="Column Name"
                placeholder="Select Column..."
                onChange={(ev, item: any) => {
                  this.SelectedColumn = item.key
                  this.LoadFindTextOptions();
                  this.FindText = [];
                  // this.setState({Coloptions : ListNames});
                }}

              ></Dropdown>

              {/* <TextField label="Column Name"

                onChange={(ev, text: any) => {
                  this.setState({ ColumnName: text });
                }} /> */}
            </div>
            <div className={styles.column}>

              <Dropdown
            
                options={this.state.FindTextOptions}
                label="Find text"
                placeholder="Select text..."
                onChange={(ev, item: any) => {
                    this.FindText = [];
                    this.SelectedFindText = item.key
                   this.getAllTextID();
                  // this.setState({Coloptions : ListNames});
                }}

              ></Dropdown>

            </div>

            <div className={styles.column}>
              <TextField label="Replace with"
               
                onChange={(ev, text: any) => {
                  this.ReplaceWithText = text;
                  this.setState({ ReplaceText: text });
                }} />
            </div>

            <div className={styles.column} style={{ marginTop: '15px', marginLeft: '220px' }}>
              <PrimaryButton onClick={() => this.buttonclick()}> {
                "Replace"}
              </PrimaryButton>
            </div>
            <Dialog
              hidden={this.state.hDialog}
              dialogContentProps={{
                type: DialogType.close,
                title: this.state.DHeader,
                subText: this.state.DMessage
              }}
              onDismiss={() => {
                this.setState({
                  hDialog: true, HideFooter: false
                });
                window.location.reload();
              }}
              modalProps={{
                isBlocking: true,
                containerClassName: 'ms-dialogMainOverride'
              }}>
              {
                this.state.HideFooter ?
                  "" :
                  <DialogFooter>
                    <PrimaryButton onClick={() => this.Replaceclick()} text="Ok" />
                    <DefaultButton onClick={() => this.setState({ hDialog: true })} text="Cancel" />
                  </DialogFooter>
              }
            </Dialog>


          </div>
        </div>
      </div>
    );
  }
}
