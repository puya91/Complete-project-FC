import * as React from 'react';
import styles from './CompleteProjectFc.module.scss';
import { ICompleteProjectFcProps } from './ICompleteProjectFcProps';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/files";
import "@pnp/sp/attachments";
import "@pnp/sp/sites";
import "@pnp/sp/site-users";
import "@pnp/sp/site-users/web";
import { useEffect, useState } from 'react';
import { getSP } from '../pnpJsConfig';
import { SPFI } from '@pnp/sp';
import { DatePicker, DefaultButton, DirectionalHint, Dropdown, FocusTrapZone, IDropdownOption, IPersonaProps, Icon, Label, Layer, MessageBar, MessageBarType, Popup, PrimaryButton, Stack, StackItem, TextField, TooltipHost, defaultDatePickerStrings } from '@fluentui/react';
import { buttonStackTokens, confirmTitleStyle, dropdown, filePickerButtonUploadStyle, inputStyle, inputStyleLarge, popupStyles, stackTokens } from '../constants/styleConstants';
import { countryOptions } from '../constants/optionsConstants';
import { IBusiness } from '../models/IBusiness';
import { FilePicker, IFilePickerResult } from '@pnp/spfx-controls-react/lib/FilePicker';
import { deleteFiles, getDocumentsUrl, onRejectOrCloseSubmit, updateListItems } from '../services/SharepointServices';
import { IDocumentData } from '../models/IDocumentData';
import { IPeoplePickerUserItem, PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { IAuthor } from '../models/IAuthor';


const CompleteProjectFc = (props: ICompleteProjectFcProps): JSX.Element => {

  const sp: SPFI = getSP();
  const [author, setAuthor] = useState<IAuthor>();
  const [currentUserId, setCurrentUserId] = useState<number>(0);
  const [businessListItems, setBusinessListItems] = useState<IBusiness[]>([]);
  const [riskTitle, setRiskTitle] = useState<string | undefined>(props.listItem.RiskTitle);
  const [peoplePickerSelectedUsers, setPeoplePickerSelectedUsers] = useState<IPeoplePickerUserItem[]>([]);
  const [business, setBusiness] = useState<string | undefined>(props.listItem.Business ? props.listItem.Business : undefined);
  const [selectedBusiness, setSelectedBusiness] = useState<string | undefined>(props.listItem.Business ? props.listItem.Business : undefined);
  const [country, setCountry] = useState<string | undefined>(props.listItem.Country ? props.listItem.Country : undefined);
  const [selectedCountry, setSelectedCountry] = useState<string | undefined>(props.listItem.Country ? props.listItem.Country : undefined);
  const [riskDate, setRiskDate] = useState<Date | undefined >(props.listItem.RiskDate ? props.listItem.RiskDate : undefined);
  const [notes, setNotes] = useState(props.listItem.AdditionalNotes);
  const [riskReport, setRiskReport] = useState<IFilePickerResult[]>();
  const [documentsToDownload, setDocumentsToDownload] = useState<IDocumentData[] | undefined>();
  const [filledObligatoryComponents, setFilledObligatoryComponents] = useState<string[]>([]);
  const [disableSendButton, setDisableSendButton] = useState<boolean>(true);
  const [isSendButtonClicked, setIsSendButtonClicked] = useState<boolean>(false);
  const [isSaveButtonClicked, setIsSaveButtonClicked] = useState<boolean>(false);
  const [disableSaveButton, setDisableSaveButton] = useState<boolean>(false);
  const [isDeletePopupVisible, setIsDeletePopupVisible] = useState<boolean>(false);
  const [isSendPopupVisible, setIsSendPopupVisible] = useState<boolean>(false);
  const [isSavePopupVisible, setIsSavePopupVisible] = useState<boolean>(false);
  const [isRejectPopupVisible, setIsRejectPopupVisible] = useState<boolean>(false);
  const [isClosePopupVisible, setIsClosePopupVisible] = useState<boolean>(false);
  const [isSendSuccess, setIsSendSuccess] = useState<boolean>(false);
  const [isSaveSuccess, setIsSaveSuccess] = useState<boolean>(false);
  const [isRejectSuccess, setIsRejectSuccess] = useState<boolean>(false);
  const [isCloseSuccess, setIsCloseSuccess] = useState<boolean>(false);
  const [areDocumentsDeleted, setAreDocumentsDeleted] = useState<boolean>(false);
  const [requestState, setRequestState] = useState<string>(props.listItem.State);



  useEffect(() => {
    setTimeout(() => {
      window.scrollTo(0, 0);
    }, 100);
  }, [isSaveSuccess, isSendSuccess, isRejectSuccess, isCloseSuccess])

  useEffect(() => {
    const fetchDocuments = async (): Promise<void> => {
      try {
        const documentsUrl = await getDocumentsUrl(props.listItem.Title);
        setDocumentsToDownload(documentsUrl);
      } catch (error) {
        console.error('Error getting Documents Url', error)
      }
    };

    if (props.listItem.ContainsDocuments !== "No") {
      fetchDocuments().catch(error => console.error('Error in useEffect for fetchDocuments() function:', error));
    }
  }, [props.listItem.Title]);

  const getListItems = async (): Promise<IBusiness[]> => {
    const items = sp.web.lists.getById(props.businessListGuid).items.orderBy('Title', true)();
    return (await items).map((item) => ({
        id: item.Id,
        title: item.Title,
        country: item.Country,
        client: item.Client
    }));
  }

  useEffect(() => {
    if(props.businessListGuid && props.businessListGuid !== '') {
      getListItems().then((items) => {
        setBusinessListItems(items);
      }).catch(error => {
        console.error('Error getting list items:', error);
      });
    }
  }, [props]);

  const getCurrentUser = async (): Promise<void> => {
    try {
      const user = await sp.web.currentUser();
      setCurrentUserId(user.Id);
    } catch (error) {
      console.error('Error getting current user:', error);
    }
  };
  
  useEffect(() => {
    getCurrentUser().catch(error => console.error('Error in useEffect for getCurrentUser() function:', error));
  }, []);

  useEffect(() => {
    const getAuthor = async (): Promise<void> => {
      try{
        const creator = await sp.web.getUserById(props.listItem.AuthorId)();

        const authorIfo = {
          Id: creator.Id,
          Name: creator.Title,
          Email: creator.Email,
          LoginName: creator.LoginName
        }

        setAuthor(authorIfo);

      } catch (error) {
        console.error('Error getting creators:', error);
      }
    }

    getAuthor().catch(error => console.error('Error getting author by getAuthor() function:', error));
  }, [props.listItem.AuthorId])

  useEffect(() => {
    const filledComponents: string[] = [];
    
    if (props.listItem.RiskTitle) {
      filledComponents.push("riskTitle");
    }
    if (props.listItem.RiskDate) {
      filledComponents.push("riskDate");
    }
    if (props.listItem.Business) {
      filledComponents.push("business");
    }
    if (props.listItem.Country) {
      filledComponents.push("country");
    }
    if (props.listItem.AssignedToPeopleId) {
      filledComponents.push("riskAssignment");
    }
  
    if (filledComponents.length > 0) {
      setFilledObligatoryComponents(prevState => [...prevState, ...filledComponents]);
    }
  }, [props.listItem]);
  

  useEffect(() => {
    if (documentsToDownload) {
      setFilledObligatoryComponents(prevState => [...prevState, "riskReport"]);
    }
  },[documentsToDownload])

  const onTitleChange = (_ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value: string | undefined, name: string): void => {
    setRiskTitle(value || '');

    if (!filledObligatoryComponents.includes(name)) {
      setFilledObligatoryComponents(prevState => [...prevState, name]);
    }
    else if(filledObligatoryComponents.includes(name) && !value) {
      setFilledObligatoryComponents(prevState => prevState.filter(item => item !== name));
    }

    setIsSendButtonClicked(false);
    setIsSaveButtonClicked(false);
    setIsSendSuccess(false);
    setIsSaveSuccess(false);
  }

  const contextWithRequiredProps = {
    ...props.context,
    absoluteUrl: props.context.pageContext.web.absoluteUrl,
    msGraphClientFactory: props.context.msGraphClientFactory,
    spHttpClient: props.context.spHttpClient,
  };

  useEffect(() => {
    const getUsers = async (): Promise<void> => {
      try {
        const userPromises = props.listItem.AssignedToPeopleId.map((Id) => {
          return sp.web.getUserById(Id)();
        });

        const users = await Promise.all(userPromises);
        const formattedUsers: IPeoplePickerUserItem[] = users.map((user) => ({
          id: user.Id.toString(),
          loginName: user.LoginName,
          imageUrl: `https://4l8x4l.sharepoint.com/_layouts/15/userphoto.aspx?accountname=${user.Email}&size=M`,
          imageInitials: user.Title.split(' ').map(name => name[0]).join('').substring(0, 2),
          text: user.Title,
          secondaryText: user.Email,
          tertiaryText: '',
          optionalText: ''
        }));
        setPeoplePickerSelectedUsers(formattedUsers);

      } catch (error) {
        console.error('Error getting users:', error);
      }
    };

      getUsers().catch(error => console.error('Error in useEffect for getUsers() function:', error));
  }, [props.listItem.AssignedToPeopleId]);

  const onPeoplePickerChange = (people: IPersonaProps[], name: string): void => {
    setPeoplePickerSelectedUsers(people as IPeoplePickerUserItem[]);
    
    if (!filledObligatoryComponents.includes(name)) {
      setFilledObligatoryComponents(prevState => [...prevState, name]);
    }
    else if(filledObligatoryComponents.includes(name) && people.length === 0) {
      setFilledObligatoryComponents(prevState => prevState.filter(item => item !== name));
    }

    setIsSendButtonClicked(false);
    setIsSaveButtonClicked(false);
    setIsSendSuccess(false);
    setIsSaveSuccess(false);
  }

  const onFormatDate = (date?: Date): string => {
    return (
      !date ? '' 
      : date.getDate() + '/' + (date.getMonth() + 1) + '/' + date.getFullYear());
  }

  const onDateChange = (value: Date | null | undefined, name: string): void => { 
    if(value) {
      setRiskDate(value);
    }

    if (!filledObligatoryComponents.includes(name)) {
      setFilledObligatoryComponents(prevState => [...prevState, name]);
    }

    setIsSendButtonClicked(false);
    setIsSaveButtonClicked(false);
    setIsSendSuccess(false);
    setIsSaveSuccess(false);
  }

  const onDropdownChange = (_event: React.FormEvent<HTMLDivElement>, item?: IDropdownOption, name?: string): void => {
    if (item && name === "business") {
      setBusiness(item.text as string);
      setSelectedBusiness(item.key as string);
    }
    else if (item && name === "country") {
      setCountry(item.text as string);
      setSelectedCountry(item.key as string);
    }
    
    if (name && !filledObligatoryComponents.includes(name)) {
      setFilledObligatoryComponents(prevState => [...prevState, name]);
    }

    setIsSendButtonClicked(false);
    setIsSaveButtonClicked(false);
    setIsSendSuccess(false);
    setIsSaveSuccess(false);
  }

  const onUpload = async (files: IFilePickerResult[], name: string): Promise<void> => {
    try {
      props.listItem.ContainsDocuments = "Yes, check RiskEventDocumentsList";
      setRiskReport(files);
      setFilledObligatoryComponents(prevState => [...prevState, name]);
      setIsSendButtonClicked(false);
      setIsSaveButtonClicked(false);
      setIsSendSuccess(false);
      setIsSaveSuccess(false);
    } catch (error){
      console.error('Error uploading documents:', error);
    }
  }

  const onDeleteConfirmation = (name: string): void => {
    props.listItem.ContainsDocuments = "No";
    setAreDocumentsDeleted(true);
    setRiskReport(undefined);
    setFilledObligatoryComponents(prevState => prevState.filter(item => item !== name));
    setIsSendButtonClicked(false);
    setIsSaveButtonClicked(false);
    setIsDeletePopupVisible(false);
  }

  const onFormValidation = (): void => {
    setIsSendButtonClicked(true);
    setIsSaveSuccess(false);

    if (filledObligatoryComponents.length === 6) {
      setIsSendPopupVisible(true);
    }
    else{
      setTimeout(() => {
        window.scrollTo(0, 0);
      }, 100);
    }
  }

  const onSubmit = async (state: string): Promise<void> => {

    const states = {
      riskTitle: riskTitle,
      selectedUsers: peoplePickerSelectedUsers.map(user => user.id),
      business: business,
      country: country,
      riskDate: riskDate,
      riskReport: riskReport,
      containsDocuments: filledObligatoryComponents.includes("riskReport") ? "Yes, check RiskEventDocumentsList" : "No",
      notes: notes,
      state: state,
    };

    if (areDocumentsDeleted) {
      await deleteFiles(props.listItem.Title);
    }

    await updateListItems(props.context, states, props.listItem.ID);
  }

  const onSaveConfirmation = async (): Promise<void>  => {
    await onSubmit("Modifying");

    setIsSavePopupVisible(false);
    setIsSaveButtonClicked(false);
    setIsSaveSuccess(true);
    setDisableSaveButton(false);
    setFilledObligatoryComponents([]);

    // Resetting all variables
    setBusiness(undefined);
    setSelectedBusiness(undefined);
    setCountry(undefined);
    setSelectedCountry(undefined);
    setRiskTitle('');
    setRiskDate(undefined);
    setPeoplePickerSelectedUsers([]);
    setNotes('');
    setRiskReport(undefined);

    setTimeout(() => {
      // window.location.reload();
      window.scrollTo(0, 0);
    }, 3000);
  }

  const onSendConfirmation = async (): Promise<void> => {
    await onSubmit("Sent");

    setIsSendPopupVisible(false);
    setDisableSendButton(true);
    setIsSendSuccess(true);
    setIsSendButtonClicked(false);
    setFilledObligatoryComponents([]);

    // Resetting all variables
    setBusiness(undefined);
    setSelectedBusiness(undefined);
    setCountry(undefined);
    setSelectedCountry(undefined);
    setRiskTitle('');
    setRiskDate(undefined);
    setPeoplePickerSelectedUsers([]);
    setNotes('');
    setRiskReport(undefined);

    setTimeout(() => {
      // window.location.reload();
      window.scrollTo(0, 0);
    }, 3000);
  }

  const onRejectConfirmation = async (): Promise<void> => {
    await onRejectOrCloseSubmit("Rejected", props.listItem.ID);

    setRequestState("Rejected");
    setIsRejectSuccess(true);

    setIsRejectPopupVisible(false);

    setTimeout(() => {
      // window.location.reload();
      window.scrollTo(0, 0);
    }, 3000);
  }

  const onCloseConfirmation = async (): Promise<void> => {
    await onRejectOrCloseSubmit("Closed", props.listItem.ID);

    setRequestState("Closed");
    setIsCloseSuccess(true);

    setIsClosePopupVisible(false);

    setTimeout(() => {
      // window.location.reload();
      window.scrollTo(0, 0);
    }, 3000);
  }

  return (
    <div className={styles.completeProjectFc} >

      {/* SUCCESS SAVE MESSAGE BAR */}
      {
        isSaveSuccess === true &&
        <Stack>
          <br /><br />
          <MessageBar messageBarType={MessageBarType.success} isMultiline={false}>
            The request has been saved successfully and the state of the request is &quot;Modifying&quot;.
          </MessageBar>
        </Stack>
      }

      {/* SUCCESS SEND MESSAGE BAR */}
      {
        isSendSuccess === true &&
        <Stack>
          <br /><br />
          <MessageBar messageBarType={MessageBarType.success} isMultiline={false}>
            The request has been sent successfully and the state of the request is &quot;Sent&quot;.
          </MessageBar>
        </Stack>
      }

      {/* SUCCESS Reject MESSAGE BAR */}
      {
        isRejectSuccess === true &&
        <Stack>
          <br /><br />
          <MessageBar messageBarType={MessageBarType.success} isMultiline={false}>
            The request has been rejected successfully and the state of the request is &quot;Rejected&quot;.
          </MessageBar>
        </Stack>
      }

      {/* SUCCESS ACCEPT AND CLOSE MESSAGE BAR */}
      {
        isCloseSuccess === true &&
        <Stack>
          <br /><br />
          <MessageBar messageBarType={MessageBarType.success} isMultiline={false}>
            The request has been accepted and closed successfully and the state of the request is &quot;Closed&quot;.
          </MessageBar>
        </Stack>
      }

      {/* ERROR MESSAGE BAR */}
      {
        (isSendButtonClicked === true && isSendSuccess === false && isSendPopupVisible === false) && 
        <Stack>
          <br /><br />
          <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
            The program has encountered a problem. You have not filled all of the obligatory fields. 
          </MessageBar>
        </Stack>
      }

      <Stack tokens={stackTokens}>
        <StackItem>
          <h1>Risk event request details</h1>
        </StackItem>
      </Stack>

      <Stack horizontal wrap tokens={stackTokens}>

        {/* RISK TITLE */}
        <StackItem>
          {
            requestState === "Sent" || 
            requestState === "Closed" || 
            requestState === "Rejected" ||
            (requestState === "Modifying" &&
             author?.Id !== currentUserId
            ) ?
              <TextField 
                borderless
                disabled
                label='Risk event Title'
                className={styles.componentStyle} 
                value={props.listItem.RiskTitle}
              />
            :
            <>
              <Label 
                className={
                  isSendButtonClicked  
                  && !filledObligatoryComponents.includes("riskTitle")  
                  ? styles.errorStyle  
                  : undefined
                }
              >
                Risk event title *
              </Label>
              <TextField 
                className={styles.componentStyle} 
                placeholder="Please write your text here"
                value={riskTitle}
                onChange={(ev, newValue) => onTitleChange(ev, newValue, "riskTitle")}
              />
            </>
          }
        </StackItem>

        {/* REQUEST STATE */}
        <StackItem>
          <TextField 
            borderless
            disabled
            label='Request state'
            className={styles.componentStyle} 
            value={requestState}
          />
        </StackItem>
      </Stack>

      <Stack tokens={stackTokens}>

        {/* ASSIGNED TO */}
        <StackItem>
          {
            requestState === "Sent" || 
            requestState === "Closed" || 
            requestState === "Rejected" ||
            (requestState === "Modifying" &&
             author?.Id !== currentUserId
            ) ?
              <>
                <Stack horizontal>
                  <Label>
                    Assigned to
                  </Label>
                  <TooltipHost 
                    content="This risk event request can be accepted or rejected by at least one assignee."
                    directionalHint={DirectionalHint.topLeftEdge}
                  >
                    <Icon iconName="Info" className={styles.peoplePickerIcon} />
                  </TooltipHost>
                </Stack>
                <Stack style={{ marginTop: 0 }}>
                  <TextField 
                    borderless
                    disabled
                    styles={inputStyleLarge}
                    value={peoplePickerSelectedUsers.map(user => user.text).join(', ')}
                  />
                </Stack>
              </>
            :
              <>
                <Stack horizontal>
                  <Label 
                    className={
                      (isSendButtonClicked || isSaveButtonClicked)
                      && !filledObligatoryComponents.includes("riskAssignment")  
                      ? styles.errorStyle  
                      : undefined
                    }
                  >
                    Assign to *
                  </Label>
                  <TooltipHost 
                    content="Write at least 3 letters for the name to appear and you can assign upto 3 people."
                    directionalHint={DirectionalHint.topLeftEdge}
                  >
                    <Icon iconName="Info" className={styles.peoplePickerIcon} />
                  </TooltipHost>
                </Stack>
                <Stack style={{ marginTop: 0 }}>
                  <PeoplePicker
                    context={contextWithRequiredProps}
                    placeholder="Please write the name here"
                    personSelectionLimit={3}
                    groupName={""} 
                    required={false}
                    disabled={false}
                    searchTextLimit={2}
                    ensureUser={true}
                    onChange={(newPeople) => onPeoplePickerChange(newPeople, "riskAssignment")}
                    principalTypes={[PrincipalType.User]}
                    resolveDelay={1000} 
                    styles={inputStyleLarge}
                    defaultSelectedUsers={peoplePickerSelectedUsers.map(user => user.text)}
                    resultFilter={(result: IPersonaProps[]) => {
                      return result.filter(person => person.id !== currentUserId.toString());
                    }}
                  />
                </Stack>
              </>
          }
          
        </StackItem>
      </Stack>

      <Stack horizontal wrap tokens={stackTokens}>

        {/* CREATED BY */}
        <StackItem>
          <TextField 
            borderless
            disabled
            label='Created by'
            className={styles.componentStyle} 
            value={author?.Name}
          />
        </StackItem>

        {/* CREATOR EMAIL */}
        <StackItem>
          <TextField 
            borderless
            disabled
            label="Creator's email"
            className={styles.componentStyle} 
            value={author?.Email}
          />
        </StackItem>
      </Stack>

      <Stack horizontal wrap tokens={stackTokens}>

        {/* CREATION DATE */}
        <StackItem>
          <TextField 
            borderless
            disabled
            label='Creation date'
            className={styles.componentStyle} 
            value={new Date(`${props.listItem.Created}`).toLocaleDateString("it-IT")}
          />
        </StackItem>
        
        {/* RISK DATE */}
        <StackItem>
          {
            requestState === "Sent" || 
            requestState === "Closed" || 
            requestState === "Rejected" ||
            (requestState === "Modifying" &&
             author?.Id !== currentUserId
            ) ?
              <TextField 
                borderless
                disabled
                label='Risk event date'
                className={styles.componentStyle} 
                value={riskDate ? new Date(`${riskDate}`).toLocaleDateString("it-IT") : undefined}
              />
            :
            <>
              <Label 
                className={
                  isSendButtonClicked  
                  && !filledObligatoryComponents.includes("riskDate")  
                  ? styles.errorStyle  
                  : undefined
                }
              >
                Risk event date *
              </Label>
              <DatePicker
                placeholder="Select a date"
                ariaLabel="Select a date"
                strings={defaultDatePickerStrings}
                styles={inputStyle}
                formatDate={onFormatDate}
                value={riskDate ? new Date(`${riskDate}`) : undefined}
                onSelectDate={(newValue) => onDateChange(newValue, "riskDate")}
              />
            </>
          }
          
        </StackItem>
      </Stack>

      <Stack horizontal wrap tokens={stackTokens}>

        {/* BUSINESS */}
        <StackItem>
          {
            requestState === "Sent" || 
            requestState === "Closed" || 
            requestState === "Rejected" ||
            (requestState === "Modifying" &&
             author?.Id !== currentUserId
            ) ?
              <TextField 
                borderless
                disabled
                label='Business'
                className={styles.componentStyle} 
                value={selectedBusiness? selectedBusiness : undefined}
              />
            :
            <>
              <Label 
                className={
                  isSendButtonClicked  
                  && !filledObligatoryComponents.includes("business")  
                  ? styles.errorStyle  
                  : undefined
                }
              >
                Business *
              </Label>
              <Dropdown
                placeholder="Select an option"
                options={businessListItems.map((item: IBusiness) => ({
                  key: item.title,
                  text: item.title
                }))}
                styles={dropdown}
                onChange={(ev, item) => onDropdownChange(ev, item, "business")}
                selectedKey={selectedBusiness}
              />
            </>
          }
        </StackItem>

        {/* COUNTRY */}
        <StackItem>
          {
            requestState === "Sent" || 
            requestState === "Closed" || 
            requestState === "Rejected" ||
            (requestState === "Modifying" &&
             author?.Id !== currentUserId
            ) ?
              <TextField 
                borderless
                disabled
                label='Country'
                className={styles.componentStyle} 
                value={selectedCountry? selectedCountry : undefined}
              />
            :
            <>
              <Label 
                className={
                  isSendButtonClicked  
                  && !filledObligatoryComponents.includes("country")  
                  ? styles.errorStyle  
                  : undefined
                }
              >
                Country *
              </Label>
              <Dropdown
                placeholder="Select an option"
                options={countryOptions}
                styles={dropdown}
                onChange={(ev, item) => onDropdownChange(ev, item, "country")}
                selectedKey={selectedCountry}
              />
            </>
          }
        </StackItem>
      </Stack>

      <Stack horizontal wrap tokens={stackTokens}>

        {/* NOTES */}
        <StackItem>
          <TextField 
            className={styles.componentStyle} 
            label="Notes" 
            placeholder={
              requestState === "Sent" || 
              requestState === "Closed" || 
              requestState === "Rejected" ||
              (requestState === "Modifying" &&
              author?.Id !== currentUserId
              ) ?
                undefined 
              :
                "Please write your text here"
            }
            multiline 
            rows={5} 
            value={notes}
            onChange={(_ev, newValue) => {
              setNotes(newValue || '');
              setIsSaveSuccess(false);
              setIsSendSuccess(false);
            }}
            disabled={
              requestState === "Sent" || 
              requestState === "Closed" || 
              requestState === "Rejected" ||
              (requestState === "Modifying" &&
               author?.Id !== currentUserId
              )
            }
          />
        </StackItem>

        {/* RISK REPORT */}
        <StackItem styles={inputStyle}>
          {
            requestState === "Sent" || 
            requestState === "Closed" || 
            requestState === "Rejected" ||
            (requestState === "Modifying" &&
              author?.Id !== currentUserId
            ) ?
              <Label>
                Risk event report
              </Label>
            :
              <Label 
                className={
                  isSendButtonClicked  
                  && !filledObligatoryComponents.includes("riskReport")  
                  ? styles.errorStyle  
                  : undefined
                }
              >
              Risk event report *
              </Label>
          }
          <Stack className={styles.reportContainerZone}>
            <Stack className={styles.reportContentZone}>
              {
                requestState === "Sent" || 
                requestState === "Closed" || 
                requestState === "Rejected" ||
                (requestState === "Modifying" &&
                 author?.Id !== currentUserId
                ) ?
                  <p className={styles.filePickerDescription}>You can download the uploaded files from the list below</p>
                : (!riskReport && (props.listItem.ContainsDocuments === "No" || isCloseSuccess || isSaveSuccess)) ?
                  <p className={styles.filePickerDescription}>Upload files from your local device using the button below</p>
                :
                  <p className={styles.filePickerDescription}>If you want you can delete your files using the button below</p>
              }
              <Stack horizontal verticalAlign="center">
                  {
                    requestState === "Sent" || 
                    requestState === "Closed" || 
                    requestState === "Rejected" ||
                    (requestState === "Modifying" &&
                     author?.Id !== currentUserId
                    ) ?
                      <ul style={ { margin: '0px', paddingLeft: '15px', overflow: 'overlay', height: '36px', width: '255px' } }>
                        {documentsToDownload?.map((doc, index) => (
                          <li key={index} style={{ fontSize: 'small' }}>
                            <a href={doc.url} target="_blank" rel="noopener noreferrer">{doc.name}</a>
                          </li>
                        ))}
                      </ul>
                    : (!riskReport && (props.listItem.ContainsDocuments === "No" || isCloseSuccess || isSaveSuccess)) ?
                      <FilePicker
                        context={props.context}
                        accepts={[ ".pdf" ]}
                        hidden={false}
                        hideLocalUploadTab={true}
                        hideLocalMultipleUploadTab={false}
                        hideOneDriveTab={true}
                        hideStockImages={true}
                        hideWebSearchTab={true}
                        hideSiteFilesTab={true}
                        hideLinkUploadTab={true}
                        hideRecentTab={true}
                        onSave={(ev) => onUpload(ev, "riskReport")}
                        buttonIconProps={{ styles: filePickerButtonUploadStyle }}
                        buttonClassName={styles.filePickerButtonUpload}
                        buttonLabel='Upload'
                      />
                    :
                      <PrimaryButton 
                        text="Delete" 
                        onClick={() => setIsDeletePopupVisible(true)}
                      />
                  }
                  <StackItem>
                    {
                    requestState === "Sent" || 
                    requestState === "Closed" || 
                    requestState === "Rejected" ||
                    (requestState === "Modifying" &&
                     author?.Id !== currentUserId
                    ) ?
                      <></>
                    : (!riskReport && (props.listItem.ContainsDocuments === "No" || isCloseSuccess || isSaveSuccess)) ?
                      <p className={styles.filePickerFormat}>pdf formats allowed</p>
                    :
                      <></>
                    }
                  </StackItem>
              </Stack>
            </Stack>
          </Stack>
        </StackItem>
      </Stack>

      {/* BOTTOM SECTION */}
      <Stack horizontal horizontalAlign="start" className={styles.bottomSection}>
        {
          requestState === "Modifying" &&
          author?.Id === currentUserId &&
            <StackItem>
              <Stack horizontal tokens={buttonStackTokens}>
                <DefaultButton 
                  text="Save to draft" 
                  disabled={disableSaveButton} 
                  onClick={() => {
                    setIsSaveButtonClicked(true);
                    setIsSendButtonClicked(false);
                    setIsSavePopupVisible(true);
                  }} 
                />
                <PrimaryButton 
                  text="Finish and Send" 
                  disabled={filledObligatoryComponents.length === 0 && disableSendButton} 
                  onClick={onFormValidation} 
                />
                <p className={styles.obligatoryField}>* Obligatory field</p>
              </Stack>
            </StackItem>
        }
        {
          requestState === "Sent" &&
          props.listItem.AssignedToPeopleId.includes(currentUserId)  &&
            <StackItem>
              <Stack horizontal tokens={buttonStackTokens}>
                <DefaultButton 
                  text="Reject" 
                  onClick={() => {
                    setIsRejectPopupVisible(true);    
                  }} 
                />
                <PrimaryButton 
                  text="Accept and close" 
                  onClick={() => {
                    setIsClosePopupVisible(true);    
                  }} 
                />
              </Stack>
            </StackItem>
        }
        
      </Stack>

      {/* POPUP SECTION */}
      {/* DELETE BUTTON POPUP */}
      {
        isDeletePopupVisible === true && (
          <Layer>
            <Popup
              className={popupStyles.root}
              role="dialog"
              onDismiss={() => setIsDeletePopupVisible(false)}
            >
              <FocusTrapZone>
                <div className={popupStyles.content}>
                  <p style={confirmTitleStyle}>Are you sure you want to delete your uploaded files?</p>
                  <p>By clicking &quot;Yes&quot; your uploaded files will be deleted and you can upload new files.</p>
                  <Stack horizontal horizontalAlign="center">
                    <DefaultButton 
                      className={styles.popupDefaultButton} 
                      onClick={() => setIsDeletePopupVisible(false)}
                    >
                      No
                    </DefaultButton>
                    <PrimaryButton 
                      className={styles.popupPrimaryButton} 
                      onClick={() => onDeleteConfirmation("riskReport")} 
                    >
                      Yes
                    </PrimaryButton>
                  </Stack>
                </div>
              </FocusTrapZone>
            </Popup>
          </Layer>
        )
      }

      {/* SAVE BUTTON POPUP */}
      {
        isSavePopupVisible === true && (
          <Layer>
            <Popup
              className={popupStyles.root}
              role="dialog"
              onDismiss={() => {
                setIsSavePopupVisible(false); 
                setIsSaveButtonClicked(false);
                setIsSendButtonClicked(false);
              }}
            >
              <FocusTrapZone>
                <div className={popupStyles.content}>
                  <p style={confirmTitleStyle}>Are you sure you want to save your session to draft?</p>
                  <p>By clicking &quot;Yes&quot; your session will be saved as draft and in &quot;Modifying&quot; state.</p>
                  <Stack horizontal horizontalAlign="center">
                    <DefaultButton 
                      className={styles.popupDefaultButton} 
                      onClick={() => {
                        setIsSavePopupVisible(false); 
                        setIsSaveButtonClicked(false);
                        setIsSendButtonClicked(false);
                      }}
                    >
                      No
                    </DefaultButton>
                    <PrimaryButton 
                      className={styles.popupPrimaryButton} 
                      onClick={() => onSaveConfirmation()} 
                    >
                      Yes
                    </PrimaryButton>
                  </Stack>
                </div>
              </FocusTrapZone>
            </Popup>
          </Layer>
        )
      }

      {/* SEND BUTTON POPUP */}
      {
        isSendPopupVisible === true && (
          <Layer>
            <Popup
              className={popupStyles.root}
              role="dialog"
              onDismiss={() => {
                setIsSendPopupVisible(false);
                setIsSendButtonClicked(false);
              }}
            >
              <FocusTrapZone>
                <div className={popupStyles.content}>
                  <p style={confirmTitleStyle}>Are you sure you want to finish your session and send it?</p>
                  <p>By clicking &quot;Yes&quot; your session will be completed and in &quot;Sent&quot; state.</p>
                  <Stack horizontal horizontalAlign="center">
                    <DefaultButton 
                      className={styles.popupDefaultButton} 
                      onClick={() => {
                        setIsSendPopupVisible(false);
                        setIsSendButtonClicked(false);
                      }}
                    >
                      No
                    </DefaultButton>
                    <PrimaryButton 
                      className={styles.popupPrimaryButton} 
                      onClick={() => onSendConfirmation()} 
                    >
                      Yes
                    </PrimaryButton>
                  </Stack>
                </div>
              </FocusTrapZone>
            </Popup>
          </Layer>
        )
      }

      {/* REJECT BUTTON POPUP */}
      {
        isRejectPopupVisible === true && (
          <Layer>
            <Popup
              className={popupStyles.root}
              role="dialog"
              onDismiss={() => {
                setIsRejectPopupVisible(false);
              }}
            >
              <FocusTrapZone>
                <div className={popupStyles.content}>
                  <p style={confirmTitleStyle}>Are you sure you want to reject this request?</p>
                  <p>By clicking &quot;Yes&quot; this request will be in &quot;Rejected&quot; state.</p>
                  <Stack horizontal horizontalAlign="center">
                    <DefaultButton 
                      className={styles.popupDefaultButton} 
                      onClick={() => {
                        setIsRejectPopupVisible(false);
                      }}
                    >
                      No
                    </DefaultButton>
                    <PrimaryButton 
                      className={styles.popupPrimaryButton} 
                      onClick={() => onRejectConfirmation()} 
                    >
                      Yes
                    </PrimaryButton>
                  </Stack>
                </div>
              </FocusTrapZone>
            </Popup>
          </Layer>
        )
      }

      {/* ACCEPT AND CLOSE BUTTON POPUP */}
      {
        isClosePopupVisible === true && (
          <Layer>
            <Popup
              className={popupStyles.root}
              role="dialog"
              onDismiss={() => {
                setIsClosePopupVisible(false);
              }}
            >
              <FocusTrapZone>
                <div className={popupStyles.content}>
                  <p style={confirmTitleStyle}>Are you sure you want to accept and close this request?</p>
                  <p>By clicking &quot;Yes&quot; this request will be in &quot;Closed&quot; state.</p>
                  <Stack horizontal horizontalAlign="center">
                    <DefaultButton 
                      className={styles.popupDefaultButton} 
                      onClick={() => {
                        setIsClosePopupVisible(false);
                      }}
                    >
                      No
                    </DefaultButton>
                    <PrimaryButton 
                      className={styles.popupPrimaryButton} 
                      onClick={() => onCloseConfirmation()} 
                    >
                      Yes
                    </PrimaryButton>
                  </Stack>
                </div>
              </FocusTrapZone>
            </Popup>
          </Layer>
        )
      }
      
    </div>
  )
}
export default CompleteProjectFc;

