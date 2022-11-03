
import * as React from "react";
import styles from "./MyTodoListWebPart.module.scss";
import getSass from "./SassLibrary.module.scss";
import { IMyTodoListWebPartProps } from "./IMyTodoListWebPartProps";
import { cloneDeep, constant, escape, update } from "@microsoft/sp-lodash-subset";
import { Item, sp } from "@pnp/sp";
import { TemplateFileType } from "@pnp/sp";
import { isArray } from "@pnp/common";
import ErrorHandlingField from './common/ErrorHandlingField';
import {
  Image,
  PrimaryButton,
  List,
  DefaultButton,
  Dialog,
  DialogType,
  Panel,
  TextField,
  Dropdown,
  IDropdownOption,
  DatePicker,
  PanelType,
  Spinner,
  SpinnerSize,
  Pivot,
  PivotItem,
  PivotLinkFormat,
  Checkbox
} from "office-ui-fabric-react";

export interface IMyTodoListWebPartState {
  isProcessing: boolean;
  showModal: boolean;
  showAddTaskModal: boolean;
  showViewModal: boolean;
  showPanel: boolean;
  items: any[];
  itemsub: any[];
  tempItem: any;
  subItem: any;
  activeItem: any;
  activeIndex: number;
  errorMsg: any;
  saveReady: boolean;
  subtasks: any[];
  editFlag: boolean;
  taskId: string;
}
const REQUIRED = [
  "Title",
  "Status",
  "DueDate",
];

const SubtaskItem = [
  {
    Name: 'Create 365 account',
    DateCompleted: 'N/A',
    Status: false,
  },
  {
    Name: 'Create playground site',
    DateCompleted: 'N/A',
    Status: false,
  },
  {
    Name: 'Assign work station',
    DateCompleted: 'N/A',
    Status: false,
  },
  {
    Name: 'Setup work environment',
    DateCompleted: 'N/A',
    Status: false,
  },
];
export default class MyTodoListWebPart extends React.Component<
  IMyTodoListWebPartProps,
  IMyTodoListWebPartState
> {
  constructor(props) {
    super(props);
    this.state = {
      isProcessing: false,
      showPanel: false,
      showModal: false,
      showAddTaskModal: false,
      showViewModal: false,
      items: [],
      itemsub: [],
      tempItem: {
        Title: '',
        Description: '',
        Status: 'Not Started',
        DueDate: new Date(),
      },
      subItem: {
        Title: '',
        Status: 'Not Started',
        DateCompleted: new Date(),
        subTestID: null,
      },
      activeItem: null,
      activeIndex: -1,
      errorMsg: {},
      saveReady: false,
      subtasks: SubtaskItem,
      editFlag: false,
      taskId: null,
    };
  }
  private _checkIsFormReady = () => {
    let { errorMsg, tempItem } = this.state;
    REQUIRED.forEach(field => {
      if (!tempItem[field] || (typeof tempItem[field] === 'string' && tempItem[field].trim() === '') ||
        (isArray(tempItem[field]) && tempItem[field].length == 0)) {
        errorMsg[field] = errorMsg[field] || 'This field is required';
      } else {
        errorMsg[field] = null;
      }
    });
    let flag = true;
    for (let k of Object.keys(errorMsg)) {
      if (errorMsg[k]) {
        flag = false;
        break;
      }
    }
    // flag = !this._checkAttachments();
    this.setState({ errorMsg, saveReady: flag });
  }
  public async componentDidMount(): Promise<void> {
    await sp.web.lists.getById('9a0a2f1e-31a2-4f42-ab52-a328502a08de').items.get()
      .then(res => {
        const items = [];
        res.forEach(item => {
          const temp = {
            ID: item.ID,
            Title: item.Title,
            Description: item.Description,
            Status: item.Status || 'Not Started',
            DueDate: item.DueDate || new Date(),
          };
          items.push(temp);
        });
        this.setState({ items });
        this.componenSubDidMount();
      });
  }
  public async componenSubDidMount(): Promise<void> {
    // await sp.web.lists.getById('c8c90709-d27b-440b-9be7-178623631746').items.filter("subTestID eq '" + item.ID + "'").get()
    //   .then(resultSet => {
    //     const itemsub = [];
    //     resultSet.forEach(item => {
    //       const temp = {
    //         ID: item.ID,
    //         Title: item.Title,
    //       };
    //       itemsub.push(temp);
    //     });
    //     this.setState({ itemsub });
    //   });
    await sp.web.lists.getById('c8c90709-d27b-440b-9be7-178623631746').items.get()
      .then(res => {
        const itemsub = [];
        res.forEach(item => {
          const temp = {
            ID: item.ID,
            Title: item.Title,

          };
          itemsub.push(temp);
        });
        this.setState({ itemsub });
      });
  }
  public render(): React.ReactElement<IMyTodoListWebPartProps> {
    const { items, itemsub, showModal, showAddTaskModal, showViewModal, showPanel, activeItem, tempItem, subItem, isProcessing, errorMsg, saveReady, } = this.state;
    return (
      <div className="ms-Grid" >
        <div className="ms-Grid-row">
          <div className="ms-Grid-col ms-sm12">
            <h1>Hello</h1>
          </div>
          <div className="ms-Grid-col ms-sm12">
            <PrimaryButton
              text="Add New"
              onClick={() => {
                // const item = {
                //   Title: "Hello",
                //   Description: "Testing",
                //   Status: "Not Started",
                //   DueDate: new Date().toLocaleString(),
                // };
                // items.push(item);
                // this.setState({ items });
                this.setState({ showPanel: true });
              }}
            />
          </div>
        </div>
        <div className={"ms-Grid-col ms-sm12 " + getSass.mt10}>
          <List
            items={cloneDeep(items)}
            onRenderCell={(
              item?: any,
              index?: number,
              isScrolling?: boolean
            ) => {
              return (
                <div
                  className={"ms-Grid-col ms-sm12 " + getSass.mb10 + " " + getSass.borderRidgeBlack1px + " " + getSass.roundedSm}
                >
                  <div className={"ms-Grid-col ms-sm8" + getSass.p15}>
                    <div className="ms-Grid-col ms-sm12">
                      ID: {item.ID}
                    </div>
                    <div className="ms-Grid-col ms-sm12">
                      Name: {item.Title}
                    </div>
                    <div className="ms-Grid-col ms-sm12">
                      Status: {item.Status}
                    </div>
                    <div className="ms-Grid-col ms-sm12">
                      DueDate: {item.DueDate.toLocaleString()}
                    </div>
                  </div>
                  <div className="ms-Grid-col ms-sm3">
                    <div className={"ms-Grid-col ms-sm4 " + getSass.m5 + " " + getSass.mtAuto}>
                      <div className="ms-Grid-col ms-sm4">

                        <DefaultButton
                          // style={{ background: '#00b7c3', borderRadius: '100%' }}
                          className={styles.btn}
                          iconProps={{ iconName: 'View' }}
                          onClick={() => {
                            this.setState({
                              showModal: true,
                              activeItem: item,
                              activeIndex: index,
                            });
                          }}
                        />
                      </div>
                    </div>
                    <div className={"ms-Grid-col ms-sm4 " + getSass.m5 + " " + getSass.mtAuto}>
                      <div className="ms-Grid-col ms-sm4">
                        <DefaultButton
                          className={styles.btn}
                          iconProps={{ iconName: 'Edit' }}
                          onClick={() => {
                            sp.web.lists.getById('c8c90709-d27b-440b-9be7-178623631746').items.filter("subTestID eq '" + item.ID + "'").get()
                              .then(resultSet => {
                                const itemsub = [];
                                resultSet.forEach(item => {
                                  const temp = {
                                    ID: item.ID,
                                    Title: item.Title,
                                  };
                                  itemsub.push(temp);
                                });
                                this.setState({ itemsub });
                              });
                            item.DueDate = new Date(item.DueDate);
                            this.setState({
                              taskId: item.ID,
                              tempItem: item,
                              showPanel: true,
                              editFlag: true,
                            }, () => {
                              console.log("state", this.state);
                            });
                          }}
                        />
                      </div>
                    </div>
                    <div className={"ms-Grid-col ms-sm4 " + getSass.m5 + " " + getSass.mtAuto}>
                      <div className="ms-Grid-col ms-sm4">
                        <DefaultButton
                          className={styles.btn}
                          iconProps={{ iconName: 'Delete' }}
                          onClick={() => {
                            this.setState({ isProcessing: true });
                            // update sharepoint list
                            sp.web.lists.getById('9a0a2f1e-31a2-4f42-ab52-a328502a08de').items.getById(item.ID).recycle().then(_ => {
                              const res = items.filter((it, num) => {
                                if (index != num) {
                                  return it;
                                }
                              });
                              this.setState({ items: cloneDeep(res), isProcessing: false });
                            });

                          }}
                          disabled={isProcessing}
                        />
                      </div>
                    </div>

                  </div>
                </div>
              );
            }}
          />
        </div>
        <Panel isOpen={showPanel}
          onDismiss={() => this.setState({ showModal: false })}
          onOuterClick={() => { }}
          type={PanelType.medium}>
          {this._handleRenderHeader()}
          <Pivot linkFormat={PivotLinkFormat.links}>
            <PivotItem headerText="Task Details">
              <div className={"ms-Grid-col sm-12 " + getSass.m10} >
                <ErrorHandlingField
                  isRequired={true}
                  label="Title"
                  errorMessage={errorMsg.Title}
                  parentClass={"ms-Grid-col ms-sm12"}
                >
                  <TextField
                    value={tempItem.Title}
                    onChanged={(newVal: string) => {
                      tempItem.Title = newVal;
                      this.setState({ tempItem }, () => {
                        this._checkIsFormReady();
                      });
                    }}
                  />
                </ErrorHandlingField>
                <div className="ms-Grid-col sm-12">
                  <TextField label="Description" value={tempItem.Description} onChanged={(newVal: string) => {
                    tempItem.Description = newVal;
                    this.setState({ tempItem }, () => {
                      this._checkIsFormReady();
                    });
                  }}
                    multiline
                    rows={6} />
                </div>
                <ErrorHandlingField
                  isRequired={true}
                  label="Status"
                  errorMessage={errorMsg.Status}>
                  <Dropdown options={[
                    { key: 'Not Started', text: 'Not Started' },
                    { key: 'Not Started', text: 'Not Started' },
                    { key: 'Not Started', text: 'Not Started' },
                    { key: 'Not Started', text: 'Not Started' },
                    { key: 'Not Started', text: 'Not Started' },
                  ]}
                    selectedKey={tempItem.Status || 'Not Started'}
                    onChanged={(option: IDropdownOption, index?: number) => {
                      tempItem.Status = option.key;
                      this.setState({ tempItem }, () => {
                        this._checkIsFormReady();
                      });
                    }}
                  />
                </ErrorHandlingField>
                <ErrorHandlingField
                  isRequired={false}
                  label="DueDate"
                  errorMessage={errorMsg.Status}>
                  <DatePicker
                    value={tempItem.DueDate}
                    onSelectDate={(date: Date) => {
                      tempItem.DueDate = date;
                      this.setState({ tempItem }, () => {
                        this._checkIsFormReady();
                      });
                    }} />
                </ErrorHandlingField>
              </div>
            </PivotItem>
            <PivotItem headerText="Subtasks">
              <div className={"ms-Grid-col sm-12 " + getSass.m10}>
                <div className="ms-Grid-col ms-sm12">
                  <PrimaryButton
                    text="Add Subtasks"
                    onClick={() => {
                      this.setState({
                        showAddTaskModal: true,
                      });
                    }}
                  />
                </div>
                {/*  List of Subtasks */}
                <div className="ms-Grid-col ms-sm12 ">
                  <List
                    items={cloneDeep(itemsub)}
                    onRenderCell={(
                      item?: any,
                      index?: number,
                      isScrolling?: boolean
                    ) => {
                      return (
                        <div
                          className={"ms-Grid-col ms-sm12 " + getSass.mb10 + " " + getSass.borderRidgeBlack1px}
                        >
                          <div className="ms-Grid-col ms-sm8">
                            <div className="ms-Grid-col ms-sm12">
                              SubTask ID: {item.ID}
                            </div>
                            <div className="ms-Grid-col ms-sm12" style={item.Status ? { textDecoration: 'line-through' } : {}}>
                              Name: {item.Title}
                            </div>
                          </div>
                          <div className="ms-Grid-col ms-sm4">
                            <div className="ms-Grid-col ms-sm12">
                              <div className="ms-Grid-col ms-sm2">
                                <Checkbox
                                  style={{ background: '$00b7c3' }}
                                  onChange={(ev, checked: boolean) => {
                                    const temp = this.state.subtasks;
                                    temp[index].Status = checked;
                                    this.setState({ subtasks: temp });
                                    console.log("Checked");
                                  }}
                                  value={item.status}
                                />

                              </div>
                            </div>
                          </div>
                        </div>
                      );
                    }}
                  />
                </div>
              </div>
            </PivotItem>
          </Pivot>
          {/* <div className="ms-Grid-col sm-12">
            <div className="ms-Grid-col sm-12">
              <div className="ms-Grid-col sm-12">
                <br /><br />
                <div className="ms-Grid-col sm-sm3">
                  <PrimaryButton
                    style={{ width: '100%', padding: '15px 10px' }}
                    text="Save"
                    onClick={async () => {
                      this.setState({ isProcessing: true });
                      // save to sharepointList
                      await sp.web.lists.getById('9a0a2f1e-31a2-4f42-ab52-a328502a08de').items.add(tempItem)
                      // query updates

                      items.push(tempItem);
                      // refresh DOM
                      this.setState({
                        items,
                        showPanel: false,
                        isProcessing: false,
                        tempItem: {
                          Title: '',
                          Description: '',
                          Status: 'Not Started',
                          DueDate: new Date()
                        }
                      });
                    }}
                    disabled={!saveReady || isProcessing}
                  />
                </div>
                <div className="ms-Grid-col sm-sm3">
                  <DefaultButton
                    style={{ width: '100%', padding: '15px 10px' }}
                    text="Cancel"
                    onClick={() => {
                      this.setState({
                        items,
                        showPanel: false,
                        tempItem: {
                          Title: '',
                          Description: '',
                          Status: 'Not Started',
                          DueDate: new Date()
                        }
                      });
                    }}
                    disabled={isProcessing} />
                </div>
              </div>
            </div></div> */}
          {this._handleRenderFooter()}
        </Panel>
        <Dialog
          hidden={!showModal}
          modalProps={{ isBlocking: false }}
          onDismiss={() => this.setState({ showModal: false, activeItem: null, activeIndex: -1 })}
          dialogContentProps={
            {
              type: DialogType.normal,
              title: 'Task Details',
              getStyles: () => {
                return {
                  main: {
                    minWidth: '75vw !important',
                    minHeight: '65vh'
                  },
                  header: {
                    height: '50px',
                  },
                  title: {
                    color: 'black'
                  },
                  topButton: {
                    padding: '10px'
                  },
                  button: {
                    color: 'black !important'
                  },
                  inner: {
                    overflowWrap: 'break-word'
                  },
                  subText: {
                    fontSize: '14px',
                    fontWeight: 'bold'
                  }
                };
              }
            }
          }
        >
          <div className="ms-Grid-col ms-sm12">
            <span style={{ textAlign: 'center' }}>
              {activeItem && (
                <div> <b>Description:</b>{activeItem.Description}</div>

              )}
            </span>
          </div>
        </Dialog>
        <Dialog
          hidden={!showViewModal}
          modalProps={{ isBlocking: false }}
          onDismiss={() => this.setState({ showViewModal: false, activeItem: null, activeIndex: -1 })}
          dialogContentProps={
            {
              type: DialogType.normal,
              title: 'Task Details',
              getStyles: () => {
                return {
                  main: {
                    minWidth: '75vw !important',
                    minHeight: '65vh'
                  },
                  header: {
                    height: '50px',
                  },
                  title: {
                    color: 'black'
                  },
                  topButton: {
                    padding: '10px'
                  },
                  button: {
                    color: 'black !important'
                  },
                  inner: {
                    overflowWrap: 'break-word'
                  },
                  subText: {
                    fontSize: '14px',
                    fontWeight: 'bold'
                  }
                };
              }
            }
          }
        >

          <div className="ms-Grid-col ms-sm12">
            <span style={{ textAlign: 'center' }}>
              {activeItem && (
                <div> <b>Description:</b>{activeItem.Description}</div>

              )}
            </span>
          </div>

        </Dialog>

        <Dialog
          hidden={!showAddTaskModal}
          modalProps={{ isBlocking: false }}
          onDismiss={() => this.setState({ showAddTaskModal: false, activeItem: null, activeIndex: -1 })}
          dialogContentProps={
            {
              type: DialogType.normal,
              title: 'Task Details',
              getStyles: () => {
                return {
                  main: {
                    minWidth: '75vw !important',
                    minHeight: '65vh'
                  },
                  header: {
                    height: '50px',
                  },
                  title: {
                    color: 'black'
                  },
                  topButton: {
                    padding: '10px'
                  },
                  button: {
                    color: 'black !important'
                  },
                  inner: {
                    overflowWrap: 'break-word'
                  },
                  subText: {
                    fontSize: '14px',
                    fontWeight: 'bold'
                  }
                };
              }
            }
          }
        >
          <div className="ms-Grid-col ms-sm12">
            <TextField
              label="Text"
              value={subItem.Title}
              onChanged={(newVal: string) => {
                subItem.Title = newVal;
                this.setState({ subItem });
              }}
            />
          </div>
          <div className="ms-Grid-row" style={{ display: "flex", marginTop: "20px" }} >
            <div className={"ms-Grid-col ms-sm12 ms-xl3 " + styles.awkwardMdtoLg3} style={{ margin: "0 15px 5px", width: "33.33%" }}>
              <DefaultButton text="cancel" />
            </div>
            <div className={"ms-Grid-col ms-sm12 ms-xl3 " + styles.awkwardMdtoLg3} style={{ margin: "0 15px 5px", width: "33.33%" }}>
              <PrimaryButton
                text="Submit"
                onClick={async () => {
                  subItem.subTestID = this.state.taskId.toString();
                  console.log("sub", subItem);
                  await sp.web.lists.getById('c8c90709-d27b-440b-9be7-178623631746').items.add(subItem)
                    .then(res => {
                      itemsub.push(subItem);
                      this.setState({
                        itemsub, showAddTaskModal: false,
                        subItem: {
                          Title: '',
                          Status: 'Not Started',
                          DateCompleted: new Date(),
                        },
                      });
                    });
                }}
              />
            </div>
          </div>
        </Dialog>
      </div >
    );
  }
  private _handleRenderHeader = () => {

    return (
      <div className={styles.siteTheme + " ms-Grid-row " + styles.panelHeaderV2} style={{ display: 'flex' }}>
        <div className={"ms-Grid-col ms-sm12 " + styles.awkwardSmtoMdHeader}>
          <div >New Todo Form</div>
        </div>
        {this.state.tempItem.status && (
          <div className={"ms-Grid-col ms-sm12 ms-xl6" + styles.awkwardSmtoMdStatus}>
            <div>{'Status: ${this.state.status}'}</div>
          </div>
        )}
      </div>
    );
  }
  private _handleRenderFooter = () => {
    const { tempItem, items, saveReady, isProcessing, editFlag } = this.state;
    return (
      <div className="ms-Grid-row" style={{ padding: "8px 0 80% 8px" }} >

        <div className="ms-Grid-row" style={{ display: "flex" }}>
          <div className={"ms-Grid-col ms-sm12 ms-xl3 " + styles.awkwardMdtoLg3} style={{ margin: "0 15px 5px", width: "33.33%" }}>
            <PrimaryButton
              style={{ width: '100%' }}
              onClick={async () => {
                this.setState({ isProcessing: true });
                if (editFlag) {
                  await sp.web.lists.getById('9a0a2f1e-31a2-4f42-ab52-a328502a08de').items.getById(tempItem.ID)
                    .update(tempItem).then(res => {
                      const temp = items.map((i, n) => {
                        if (i.ID == tempItem.ID) {
                          return tempItem;
                        }
                        else {
                          return i;
                        }
                      });
                      this.setState({
                        items, showPanel: false, editFlag: false, isProcessing: false,
                        tempItem: {
                          Title: '',
                          Description: '',
                          Status: 'Not Started',
                          DueDate: new Date()
                        }
                      });
                    });
                } else {
                  await sp.web.lists.getById('9a0a2f1e-31a2-4f42-ab52-a328502a08de').items.add(tempItem)
                    .then(res => {
                      tempItem.ID = res.data.ID;
                      items.push(tempItem);

                      this.setState({
                        items, showPanel: false, editFlag: false, isProcessing: false,
                        tempItem: {
                          Title: '',
                          Description: '',
                          Status: 'Not Started',
                          DueDate: new Date()
                        }
                      });
                    });
                }

              }}
              disabled={!saveReady || isProcessing}
            >
              Save
              {isProcessing && (
                <Spinner
                  size={SpinnerSize.small}
                  style={{ marginLeft: "5px" }}
                />
              )}
            </PrimaryButton>

          </div>
          <div className={"ms-Grid-col ms-sm12 ms-xl3 " + styles.awkwardMdtoLg3} style={{ width: "33.33%" }}>
            <DefaultButton
              style={{ width: '100%' }}
              text="Cancel"
              onClick={() => {

                this.setState({
                  showPanel: false, editFlag: false, tempItem: {
                    tempItem: {
                      Title: '',
                      Description: '',
                      Status: 'Not Started',
                      DueDate: new Date()
                    }
                  }
                });
              }}
              disabled={isProcessing}
            />
          </div>
        </div>
      </div>
    );
  }

}
