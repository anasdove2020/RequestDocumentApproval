// /*eslint @typescript-eslint/no-explicit-any : 0*/
// /*eslint prefer-const: 0*/
// /*eslint @typescript-eslint/no-unused-vars:0*/

// import * as React from 'react';

// import { File, ViewType, MgtTemplateProps } from "@microsoft/mgt-react";
// import { IfollowDocumentProps } from "./IfollowDocumentProps";
// import { IfollowDocumentState } from "./IfollowDocumentState";
// import { DefaultButton, IconButton, PrimaryButton } from "@fluentui/react/lib/Button";
// import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
// import { IStackTokens, Stack } from '@fluentui/react/lib/Stack';
// import {
//     isListProvisioned, isfollowedcheck,
//     deleteListItem, UpdateFollowingList,
//     getSharingLinkForFile, createFollowingListItem, followDocumentGraphAPI, stopfollowingDocumentGraphAPI, getSPSiteID
// } from "../../Services/GraphAPI";
// import {
//     TextField, DialogContent
// } from "@fluentui/react";
// import { Dropdown, IDropdownStyles, IDropdownOption } from '@fluentui/react/lib/Dropdown';
// import styles from './followDocument.module.scss';
// import { CategoryDialog } from '../CategoryDialog/CategoryDialog';
// import { defaultCategoryName } from '../../util/Constant';
// import { IFileProperties } from '../../FollowDocuments/FollowDocumentsCommandSet';

// let options: IDropdownOption[] = [];

// const dropdownStyles: Partial<IDropdownStyles> = {
//     dropdown: { width: 400 }
// };

// let siteID: any = "";
// let categoryListID: any = "";
// let myFollowingListID: any = "";
// let notesInput: any = "";

// /******Single Category Follow******/
// //let selectedCategoryOption: any = "";
// let listItemID: any = "";

// /******Multiple Category Follow******/
// let selectedCategoryOption: any[] = [];
// //let listItemID: any[] = [];
// let listData: any[] = [];


// export class RequestApproval extends React.Component<IRequestApprovalProps, IRequestApprovalState> {

//     constructor(props: IRequestApprovalProps | Readonly<IRequestApprovalProps>) {
//         super(props);

//         this.state = {
//             dataLoading: true,
//             fileInfo: this.props.fileInfo,
//             context: this.props.context,
//             // isModalOpen: true,
//             enableCategoryAddition: false,
//             dropdownOptions: options,
//             currentlyEditingCategory: null,
//             defaultSelectedCategory: [],
//             defaultNotes: "",
//             showCategoryDialog: false
//         };


//         isListProvisioned(this.props).then((followingCatagoryData) => {
//             if (followingCatagoryData !== undefined && followingCatagoryData.length > 0 &&
//                 followingCatagoryData[0] && followingCatagoryData[0].value.length > 0
//                 && followingCatagoryData[1] && followingCatagoryData[2] && followingCatagoryData[3]
//             ) {
//                 // Show Following Category list

//                 //reset variable
//                 options = [];

//                 //add updated list
//                 for (let category of followingCatagoryData[0].value) {
//                     options.push({ key: category.fields.id, text: category.fields.Title });
//                 }

//                 //reset vars
//                 selectedCategoryOption = [];
//                 listItemID = [];
//                 //end

//                 siteID = followingCatagoryData[1];
//                 categoryListID = followingCatagoryData[2];
//                 myFollowingListID = followingCatagoryData[3];

//                 this.isfollowed(this.props.fileInfo);

//             } else {
//                 // Provision Lists - Following Catagory
//                 alert('There seems to be intermittent error.\nPlease refresh and try again.');

//             }
//         });
//     }

//     /**
//      * Method to get the current following status of the selected document
//      * @param fileItemInfo 
//      */
//     private async isfollowed(fileItemInfo: IFileProperties[]) {
//         const state = await isfollowedcheck(siteID, myFollowingListID, this.props, fileItemInfo);

//         /******************Single category follow*********************/
//         //#region Single category follow
//         // const graphData: any = await stopfollowingDocumentAPI(siteID, myFollowingListID, listItemID, fileInfo, this.props);
//         // if (graphData === undefined) {
//         //     selectedCategoryOption = options.filter(item => { if (item.text == defaultCategoryName) { return item; } })[0];
//         //     notesInput = "";
//         //     listItemID = "";

//         //     this.setState({
//         //         followStatus: false,
//         //         defaultNotes: "",
//         //         defaultSelectedCategory: selectedCategoryOption
//         //     });
//         // }
//         // if (state[0] && state[0].length > 0 && state[1] && state[1].length > 0
//         //     && state[1][0] && state[1][0].fields
//         //     /* 
//         //     //commented due to delete category - uncategorised handling
//         //     && state[1][0].fields.Category
//         //     */
//         // ) {

//         //     if (state[1][0].fields.Category) {

//         //         let filteredCatArr = options.filter(item => {
//         //             if (item.text == state[1][0].fields.Category) {
//         //                 return item;
//         //             }
//         //         });

//         //         if (filteredCatArr.length > 0) {
//         //             selectedCategoryOption = filteredCatArr[0];
//         //         }
//         //     } else {
//         //         //Uncategorsied document
//         //         selectedCategoryOption = "";//options.filter(item => { if (item.text == defaultCategoryName) { return item; } })[0];
//         //     }


//         //     //listItemID = state[1] && state[1][0] && state[1][0].fields && state[1][0].fields.id ? state[1][0].fields.id : "";
//         //     notesInput = state[1] && state[1][0] && state[1][0].fields && state[1][0].fields.Note ? state[1][0].fields.Note : "";

//         //     this.setState({
//         //         dataLoading: false,
//         //         followStatus: true,
//         //         isModalOpen: false,
//         //         defaultSelectedCategory: selectedCategoryOption,
//         //         defaultNotes: notesInput,
//         //         dropdownOptions: options
//         //     });
//         // } else {

//         //     selectedCategoryOption = options.filter(item => { if (item.text == defaultCategoryName) { return item; } })[0];
//         //     listItemID = state[1] && state[1][0] && state[1][0].fields && state[1][0].fields.id ? state[1][0].fields.id : "";
//         //     notesInput = state[1] && state[1][0] && state[1][0].fields && state[1][0].fields.Note ? state[1][0].fields.Note : "";

//         //     this.setState({
//         //         dataLoading: false,
//         //         followStatus: false,
//         //         defaultSelectedCategory: selectedCategoryOption,
//         //         defaultNotes: notesInput,
//         //         dropdownOptions: options
//         //     });
//         // }

//         //#endregion

//         /******************Multiple category follow*********************/

//         //#region 
//         if (state[0] && state[0].length > 0 && state[1] && state[1].length > 0
//             && state[1][0]
//             /* 
//             //commented due to delete category - uncategorised handling
//             && state[1][0].fields.Category
//             */
//         ) {

//             //Create array of all followed category ID and list Item ID
//             let followListItem = state[1][0];

//             //Create array of list items data
//             listData.push(followListItem.fields);

//             if (followListItem.fields && followListItem.fields.Category) {
//                 followListItem.fields.Category.map((eachCategoryItem: { LookupId: { toString: () => any; }; }) => {
//                     if (eachCategoryItem.LookupId) {
//                         selectedCategoryOption.push(eachCategoryItem.LookupId.toString());
//                     }
//                 });

//             }

//             listItemID = state[1] && state[1][0] && state[1][0].fields && state[1][0].fields.id ? state[1][0].fields.id : "";
//             notesInput = state[1] && state[1][0] && state[1][0].fields && state[1][0].fields.Note ? state[1][0].fields.Note : "";

//             this.setState({
//                 dataLoading: false,
//                 followStatus: true,
//                 defaultSelectedCategory: selectedCategoryOption,
//                 defaultNotes: notesInput,
//                 dropdownOptions: options
//             });
//         } else {

//             selectedCategoryOption.push(options.filter(item => { if (item.text == defaultCategoryName) { return item; } })[0].key);
//             listItemID = "";
//             notesInput = state[1] && state[1][0] && state[1][0].fields && state[1][0].fields.Note ? state[1][0].fields.Note : "";

//             this.setState({
//                 dataLoading: false,
//                 followStatus: false,
//                 defaultSelectedCategory: selectedCategoryOption,
//                 defaultNotes: notesInput,
//                 dropdownOptions: options
//             });
//         }
//         //#endregion
//     }

//     /**
//      * Close function to close the modal
//      */
//     private closeFollowModal = () => {

//         //Reset variables to blank
//         // selectedCategoryOption = "";
//         // listItemID = "";

//         selectedCategoryOption = [];
//         listItemID = [];
//         this.props.close();

//         // this.setState({
//         //     isModalOpen: false
//         // });
//     }

//     /**
//      * Method to handle category dropdown changes
//      * @param event 
//      * @param option 
//      * @param index 
//      */
//     private handleDropdownChange(event: { target: { id: any; }; }, option: { selected: any; key: string; }, index: any) {

//         /************************************Multiple Category follow************************/

//         if (option) {
//             if (option.selected && option.key && selectedCategoryOption.indexOf(option.key) === -1) {
//                 selectedCategoryOption.push(option.key as string);
//             } else {
//                 let tempSelCatArr = selectedCategoryOption.filter(item => {
//                     if (item !== option.key) {
//                         return item;
//                     }
//                 });

//                 //reassign current selected categories
//                 selectedCategoryOption = tempSelCatArr;

//             }
//         }

//         this.setState({
//             defaultSelectedCategory: selectedCategoryOption
//         });
//     }

//     /**
//      * Method to store the input for Note field
//      * @param event trigger event, contains input value
//      */
//     private handleNotesInputChange(event: { target: { value: any; }; }) {
//         notesInput = event.target.value;

//         this.setState({
//             defaultNotes: notesInput
//         });
//     }

//     /**
//        * method called when new bulk files selected
//        * @param nextProps updated props
//        */
//     public componentWillReceiveProps(nextProps: { fileInfo: IFileProperties[]; }) {

//         if (nextProps.fileInfo !== this.props.fileInfo) {
//             //Reset variables to blank - single category follow
//             //selectedCategoryOption = "";
//             listItemID = "";

//             selectedCategoryOption = [];
//             //listItemID = [];

//             this.setState({
//                 dataLoading: true,
//                 fileInfo: nextProps.fileInfo
//             });
//             this.isfollowed(nextProps.fileInfo);
//         } else {
//             this.setState({
//                 dataLoading: false
//             });

//         }
//     }

//     /**
//     * Callback functions from child component - Category Dialog
//     * @param updatedOptions - updated category dropdown options 
//     */
//     private updateCategoryDropdown(updatedOptions: any) {
//         this.setState({
//             dropdownOptions: updatedOptions,
//             showCategoryDialog: false
//         });
//     }

//     /**
//      * Method to toggle category dialog show/hide
//      */
//     private toggleCategoryModalView() {
//         this.setState({
//             showCategoryDialog: !this.state.showCategoryDialog
//         });
//     }

//     public render(): React.ReactElement<IfollowDocumentProps> {

//         const { fileInfo, /*isModalOpen,*/ dataLoading } = this.state;

//         /**
//          * Function call to follow document
//          */
//         const followDocument = async () => {

//             this.setState({
//                 dataLoading: true
//             });

//             const getSharingLink = await getSharingLinkForFile(fileInfo, this.props);

//             if (getSharingLink != null) {
//                 fileInfo[0].fileUrl = getSharingLink;
//             }

//             //single item with multiple categories linked approach

//             const listData: any = await createFollowingListItem(
//                 siteID,
//                 myFollowingListID,
//                 fileInfo,
//                 selectedCategoryOption,
//                 notesInput,
//                 this.props.context.pageContext.site.id.toString(),
//                 this.props
//             );

//             // if (graphData[0] && graphData[1] && graphData[0].followed !== undefined && graphData[1].id) {
//             if (listData && listData.id) {
//                 listItemID = listData.id;
//             }

//             if (listItemID.length > 0 && listItemID !== "") {

//                 const graphData: any = await followDocumentGraphAPI(fileInfo, fileInfo[0].DriveId, fileInfo[0].ItemID, this.props);

//                 if (graphData && graphData.followed) {
//                     this.setState({
//                         followStatus: true,
//                         dataLoading: false
//                     });

//                     this.closeFollowModal();
//                 }
//             } else {
//                 alert("An error has occured while following this document.\nPlease refresh and try again.");
//                 this.setState({
//                     dataLoading: false
//                 });
//             }

//         };


//         /**
//          * Function call to update function on button click, to update the following data
//          */
//         const updateFollowingDocumentInfo = async () => {

//             //#region one list item with multiple tagged categories approach

//             this.setState({
//                 dataLoading: true
//             });

//             if (selectedCategoryOption == undefined || selectedCategoryOption.length < 1
//                 /*|| selectedCategoryOption == "" || selectedCategoryOption.key == undefined || selectedCategoryOption.key == ""*/
//             ) {
//                 //If blank or undefined, then assign the first dropdown option as category
//                 selectedCategoryOption.push(options[0]);
//             }

//             let updateListData: any = await UpdateFollowingList(
//                 siteID,
//                 myFollowingListID,
//                 this.props,
//                 listItemID,
//                 selectedCategoryOption,
//                 notesInput
//             );

//             if (updateListData.fields !== undefined) {
//                 //alert("Followed item has been updated");
//                 this.setState({
//                     dataLoading: false
//                 });

//                 this.closeFollowModal();
//             }

//             //#endregion

//         };

//         /**
//          * Function call to stop following document on button click
//          */
//         const stopfollowingDocument = async () => {
//             /******************Multiple category follow*********************/

//             //#region one list item with multiple tagged categories approach
//             this.setState({
//                 dataLoading: true,
//             });

//             const deletedListItem: any = await deleteListItem(siteID, myFollowingListID, listItemID, this.props);

//             if (deletedListItem === 204) {
//                 const graphData: any = await stopfollowingDocumentGraphAPI(fileInfo[0].DriveId, fileInfo[0].ItemID, this.props);

//                 if (graphData === 204) {
//                     //clear existing selection
//                     selectedCategoryOption = [];

//                     //assign default value

//                     let tempOpt = options.filter(item => {
//                         if (item.text == defaultCategoryName) {
//                             return item;
//                         }
//                     })[0];

//                     if (tempOpt.key) {
//                         selectedCategoryOption.push(
//                             tempOpt.key.toString()
//                         );
//                     }

//                     notesInput = "";
//                     listItemID = "";

//                     this.setState({
//                         dataLoading: false,
//                         followStatus: false,
//                         defaultNotes: "",
//                         defaultSelectedCategory: selectedCategoryOption
//                     });

//                     this.closeFollowModal();
//                 }
//             }
//             //#endregion
//         };

//         /**
//          * Render Loading spinner
//          * @param props 
//          * @returns 
//          */
//         const Loading = (props: MgtTemplateProps) => {
//             return <Spinner size={SpinnerSize.large} />;
//         };

//         // const dialogProps = { showCloseButton: true, title: "Follow Status" }

//         //#region Theme consts
//         const CategoryListStackTokens: IStackTokens = { childrenGap: 5 };
//         const FollowingStackTokens: IStackTokens = { childrenGap: 20 };
//         const ButtonStackTokens: IStackTokens = { childrenGap: 20, padding: 20 };

//         //#endregion

//         return (
//             //#region 
//             <DialogContent
//                 title={dataLoading ? "Bookmark" : (this.state.followStatus ? "Edit Bookmark" : "Add Bookmark")}
//                 showCloseButton={true}
//                 onDismiss={this.closeFollowModal}
//             >
//                 {/* Show loading */}
//                 {(dataLoading) && <div><Spinner size={SpinnerSize.large} /></div>}
//                 {
//                     (!dataLoading) &&
//                     <div>
//                         <Stack horizontal tokens={CategoryListStackTokens} verticalAlign="end">
//                             <Dropdown
//                                 id="CategoryDropdown"
//                                 placeholder="Select a category"
//                                 label="* Category:"
//                                 /******Single Category Follow******/
//                                 //selectedKey={this.state.defaultSelectedCategory.key}
//                                 options={this.state.dropdownOptions}//{options}
//                                 styles={dropdownStyles}
//                                 onChange={this.handleDropdownChange.bind(this)}
//                                 /******Multiple Category Follow******/
//                                 multiSelect={true}
//                                 defaultSelectedKeys={this.state.defaultSelectedCategory}
//                             />
//                             <IconButton
//                                 iconProps={{ iconName: "Edit" }}
//                                 title="Edit Categories"
//                                 ariaLabel="Edit Categories"
//                                 styles={{ root: { marginBottom: 0 } }}
//                                 onClick={this.toggleCategoryModalView.bind(this)}
//                                 disabled={!this.state.dropdownOptions || this.state.dropdownOptions.length == 0}
//                             />
//                         </Stack> <br />

//                         <Stack horizontal tokens={FollowingStackTokens} verticalAlign="end"
//                             className={styles.fileContainer}
//                             title={this.state.fileInfo[0].fileLeafRef}
//                         >
//                             <File
//                                 view={ViewType.threelines}
//                                 driveId={this.state.fileInfo[0].DriveId}
//                                 itemId={this.state.fileInfo[0].ItemID}
//                             // driveId={this.props.fileInfo[0].DriveId} itemId={this.props.fileInfo[0].ItemID}
//                             >
//                                 <Loading template="loading"></Loading>
//                             </File>
//                         </Stack>

//                         <TextField
//                             id="followingNoteInput"
//                             placeholder="Please enter note"
//                             label="Note:"
//                             multiline={true}
//                             onChange={this.handleNotesInputChange.bind(this)}
//                             className={styles.addNote}
//                             value={this.state.defaultNotes}
//                         />

//                         <Stack horizontal tokens={ButtonStackTokens} reversed>
//                             <DefaultButton text="Cancel" onClick={this.closeFollowModal.bind(this)} allowDisabledFocus />

//                             {this.state.followStatus ? (
//                                 <PrimaryButton
//                                     text="Remove"
//                                     onClick={stopfollowingDocument}
//                                     allowDisabledFocus
//                                 />
//                             ) : (
//                                 <PrimaryButton
//                                     text="Save"
//                                     onClick={followDocument}
//                                     allowDisabledFocus
//                                     disabled={selectedCategoryOption.length === 0}
//                                 />
//                             )}

//                             {this.state.followStatus
//                                 &&
//                                 <PrimaryButton
//                                     text="Save"
//                                     onClick={updateFollowingDocumentInfo}
//                                     allowDisabledFocus
//                                     disabled={selectedCategoryOption.length === 0}
//                                 />
//                             }

//                         </Stack>
//                     </div>
//                 }

//                 {
//                     this.state.showCategoryDialog &&
//                     <CategoryDialog
//                         context={this.props.context}
//                         showCategoryModal={this.state.showCategoryDialog}
//                         categoryItems={this.state.dropdownOptions}
//                         siteID={siteID}
//                         categoryListID={categoryListID}
//                         updatedCategoryItems={this.updateCategoryDropdown.bind(this)}
//                     />
//                 }

//             </DialogContent >
//             //#endregion
//         );
//     }
// }
