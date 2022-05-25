import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { DatePicker, defaultDatePickerStrings, IDatePicker } from '@fluentui/react/lib/DatePicker';
import { Dropdown, DropdownMenuItemType, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { Link } from '@fluentui/react/lib/Link';
import { IStackProps, IStackStyles, Stack } from '@fluentui/react/lib/Stack';
import { Text } from '@fluentui/react/lib/Text';
import { TextField } from '@fluentui/react/lib/TextField';
import { PartialTheme, ThemeProvider } from '@fluentui/react/lib/Theme';
import { sp } from "@pnp/sp";
import '@pnp/sp/items';
import '@pnp/sp/lists';
import '@pnp/sp/webs';
import { ListItemAttachments } from '@pnp/spfx-controls-react/lib/ListItemAttachments';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import * as React from 'react';
import { useEffect, useRef, useState } from 'react';
import { DeepMap, FieldError, useForm } from 'react-hook-form';
import { nameof } from '../utils/utils';
import { ControlledDropdown } from './controlledInputs/ControlledDropdown';
import { ControlledTextField } from './controlledInputs/ControlledTextField';
import { ISaccProps } from './ISaccProps';



// Custom ThemeProvider Settings for RH style theming
const rhTheme: PartialTheme = {
  palette: {
    themePrimary: '#9f1c37',
    themeLighterAlt: '#fbf3f5',
    themeLighter: '#efd0d6',
    themeLight: '#e2aab5',
    themeTertiary: '#c56477',
    themeSecondary: '#aa2f48',
    themeDarkAlt: '#8e1a31',
    themeDark: '#781629',
    themeDarker: '#59101e',
    neutralLighterAlt: '#faf9f8',
    neutralLighter: '#f3f2f1',
    neutralLight: '#edebe9',
    neutralQuaternaryAlt: '#e1dfdd',
    neutralQuaternary: '#d0d0d0',
    neutralTertiaryAlt: '#c8c6c4',
    neutralTertiary: '#bbb0ae',
    neutralSecondary: '#a59896',
    neutralPrimaryAlt: '#90827f',
    neutralPrimary: '#382e2c',
    neutralDark: '#645654',
    black: '#4e423f',
    white: '#ffffff',
  }
};

// Set Stack Config and Stack Styles (grid layout for input controls)
const stackTokens = { childrenGap: 100 };
const iconProps = { iconName: 'Calendar' };
const stackStyles: Partial<IStackStyles> = { root: { width: 990 } };
const columnProps: Partial<IStackProps> = {
  tokens: { childrenGap: 15 },
  styles: { root: { width: 400 } },
};


const emailDlOptions = [
  { key: 'HQPSABGCSETUP@roberthalf.com;', text: 'HQP SA BGC SETUP' },
  { key: 'sacompliance@roberthalf.com;', text: ' RHI SA Compliance ' },
  { key: 'ma.billing.profiles@roberthalf.com;', text: 'HQP MA Billing Profiles' },
  { key: 'RHISAContentData@roberthalf.com;', text: ' RHI SA Content & Data Systems ' },
  { key: 'sa.preplacementchicago@roberthalf.com;', text: ' RHI SA Preplacement-Chicago ' },
  { key: 'sa.preplacementcolumbus@roberthalf.com;', text: ' RHI SA Preplacement-Columbus ' },
  { key: 'sateama@roberthalf.com;', text: ' RHI SA Services San Ramon A ' },
  { key: 'sateamc@roberthalf.com;', text: ' RHI SA Services Columbus ' },
  { key: 'sateame@roberthalf.com;', text: ' RHI SA Services San Ramon B ' },
  { key: 'sateamf@roberthalf.com;', text: ' RHI SA Services Chicago ' },
  { key: 'sa.preplacementsanramon@roberthalf.com;', text: ' RHI SA Preplacement San Ramon ' },
  { key: 'billingtempUSEast@roberthalf.com;', text: ' Billing Temp US East ' },
  { key: 'billingtempuswest@roberthalf.com;', text: ' Billing Temp US West ' },
  { key: 'sateamg@roberthalf.com;', text: ' RHI SA Services San Ramon G ' },
  { key: 'sateamhc@roberthalf.com;', text: ' RHI SA Services San Ramon HC ' },
  { key: 'compucom@roberthalf.com;', text: ' RHI CompuCom ' },
  { key: 'usbank@roberthalf.com;', text: ' HQP US BANK National Account ' },
  { key: 'jpmc@roberthalf.com;', text: ' HQP JPMC National Account ' },
  { key: 'wellsfargo@roberthalf.com;', text: ' HQP Wells Fargo National Account ' },
  { key: 'bankingservices@roberthalf.com;', text: ' HQP SA Banking Services ' },
  { key: 'bankingservicescanada@roberthalf.com;', text: ' HQP Banking Services Canada ' },
];

// Options that are inserted into the subRequestType dropdown component
const subRequestTypeOptions = [
  { key: '--- BGC Updates', text: '--- BGC Updates', itemType: DropdownMenuItemType.Header },
  { key: 'Change Vendor - FADV', text: 'Change Vendor - FADV', dataRequestType: 'Background Check Update' },
  { key: 'Change Vendor - Non-FADV', text: 'Change Vendor - Non-FADV', dataRequestType: 'Background Check Update' },
  { key: 'Scope change - BGC/Drug/Health', text: 'Scope change - BGC/Drug/Health', dataRequestType: 'Background Check Update' },
  { key: 'Adjudication Update', text: 'Adjudication Update', dataRequestType: 'Background Check Update' },
  { key: 'Remove requirements', text: 'Remove requirements', dataRequestType: 'Background Check Update' },
  { key: 'BGC - Other', text: 'BGC - Other', dataRequestType: 'Background Check Update' },
  { key: '--- Pre-placement', text: '--- Pre-placement', itemType: DropdownMenuItemType.Header },
  { key: 'New documents', text: 'New documents', dataRequestType: 'Pre-Placement' },
  { key: 'Updated documents', text: 'Updated documents', dataRequestType: 'Pre-Placement' },
  { key: 'Remove documents', text: 'Remove documents', dataRequestType: 'Pre-Placement' },
  { key: 'New Supplier Manual', text: 'New Supplier Manual', dataRequestType: 'Pre-Placement' },
  { key: 'Update Supplier Manual', text: 'Update Supplier Manual', dataRequestType: 'Pre-Placement' },
  { key: 'SOW - Add', text: 'SOW - Add', dataRequestType: 'Pre-Placement' },
  { key: 'SOW - Update', text: 'SOW - Update', dataRequestType: 'Pre-Placement' },
  { key: 'SOW - Remove', text: 'SOW - Remove', dataRequestType: 'Pre-Placement' },
  { key: 'Pre-placement - Other', text: 'Pre-placement - Other', dataRequestType: 'Pre-Placement' },
  { key: '--- Marketing', text: '--- Marketing', itemType: DropdownMenuItemType.Header },
  { key: 'Marketing Guidelines', text: 'Marketing Guidelines', dataRequestType: 'Marketing Update' },
  { key: '1099/C2C', text: '1099/C2C', dataRequestType: 'Marketing Update' },
  { key: 'Marketing - Other', text: 'Marketing - Other', dataRequestType: 'Marketing Update' },
  { key: '--- Billing Update', text: '--- Billing Update', itemType: DropdownMenuItemType.Header },
  { key: 'OT/DT', text: 'OT/DT', dataRequestType: 'Billing Change' },
  { key: 'Bill Spec', text: 'Bill Spec', dataRequestType: 'Billing Change' },
  { key: 'Company Record/Job Order', text: 'Company Record/Job Order', dataRequestType: 'Billing Change' },
  { key: 'Billing Update - Other', text: 'Billing Update - Other', dataRequestType: 'Billing Change' },
  { key: '--- Site Update', text: '--- Site Update', itemType: DropdownMenuItemType.Header },
  { key: 'Pricing - Pricing Language', text: 'Pricing - Pricing Language', dataRequestType: 'Site Update' },
  { key: 'MSP Fee - MSP Fee Language', text: 'MSP Fee - MSP Fee Language', dataRequestType: 'Site Update' },
  { key: 'Time/Expense', text: 'Time/Expense', dataRequestType: 'Site Update' },
  { key: 'Definitive List (SALT)', text: 'Definitive List (SALT)', dataRequestType: 'Site Update' },
  { key: 'Order Process', text: 'Order Process', dataRequestType: 'Site Update' },
  { key: 'Site Update - Other', text: 'Site Update - Other', dataRequestType: 'Site Update' },
  { key: '--- Transitions', text: '--- Transitions', itemType: DropdownMenuItemType.Header },
  { key: 'Account Transition - SA to RA', text: 'Account Transition - SA to RA', dataRequestType: 'Transitions' },
  { key: 'Account Transition - RA to SA', text: 'Account Transition - RA to SA', dataRequestType: 'Transitions' },
  { key: '--- CIM Update', text: '--- CIM Update', itemType: DropdownMenuItemType.Header },
  { key: 'CIM - New site', text: 'CIM - New site', dataRequestType: 'CIM Update' },
  { key: 'CIM - Other', text: 'CIM - Other', dataRequestType: 'CIM Update' },

];

// TYPE declarations for the react-hook-form controlled inputs
type Form = {
  submitFlag: boolean;
  vpOnAccount: string;
  clientName: string;
  acctDesignation: string;
  clientServiceTeam: string;
  emailDL: string;
  subRequestType: string;
  requestType: string;
  otherRequestType: string;
  submittedBy: string;
  requestDescription: string;
  origination: string;
  otherOrigination: string;
  clientAudit: string;
  pendingReqApproval: string;
  candidateStartDate: string;
  notificationList: string;
  additionalInfo: string;
  additionalInfoSummary: string;
  status: string;
  cimReviewer: string;
  legalReviewer: string;
  saReviewer: string;
  legalReviewerResponse: string;
  cimReviewerResponse: string;
  internalNotes: string;
  internalNoteSummary: string;
};





// JSX Component that renders on screen
const Sacc: React.FC<ISaccProps> = (props) => {
  // const [itemId, setItemId] = useState<string>('');
  // const [itemState, setItemState] = useState<string>('');
  // const [submitFlag, setSubmitFlag] = useState<boolean>(false);
  const [userInfo, setUserInfo] = useState<any>({ Title: '' });
  const [emailDlKeys, setEmailDlKeys] = useState<string[]>([]);
  const [clientNames, setClientNames] = useState<any[]>([]);
  const [dateValue, setDateValue] = useState<Date | undefined>();
  const [pendingReqApprovalValue, setPendingReqApprovalValue] = useState<IDropdownOption>({ key: '', text: '' });
  const [validFormData, setValidFormData] = useState<Form>();
  const [validationError, setValidationError] = useState<DeepMap<Form, FieldError>>();

  let listItemAttachmentsComponentReference = useRef<ListItemAttachments>();
  let datePickerRef = React.useRef<IDatePicker>(null);




  let vars = [];
  let hash = [];
  let q = document.URL.split('?')[1];
  let itemIdVal = '';
  let itemStateVal = '';
  let notificationListItems = [];
  const user = sp.web.currentUser();

  // hash table setup for query string parameters
  if ('undefined' !== typeof q) {
    var p = q.split('&');

    p.forEach((val) => {
      hash = val.split('=');
      vars[hash[0]] = hash[1];
    });

    itemIdVal = vars['dsid'];
    itemStateVal = vars['stat'];

    console.log(itemIdVal);
    // setItemId(itemIdVal);
    console.log(itemStateVal);
    // setItemState(itemStateVal);
  }

  console.log('pendingReqApprovalValue: ' + pendingReqApprovalValue);

  // react hook form init
  const { handleSubmit, errors, control, register, unregister, setValue, getValues, watch } = useForm<Form, any>({
    reValidateMode: "onChange",
    shouldFocusError: true,
    mode: "all"
  });
  // setup react-hook-form "watches" to monitor component values and act on them (useful for conditional rendering)
  const watchOrigination = watch('origination');
  const watchPendingReqApproval = watch('pendingReqApproval');


  // people picker logic
  const _getPeoplePickerItems = (items) => {
    let userEmails = '';
    console.log('Items:', items);
    notificationListItems = [];
    items.forEach(val => {
      notificationListItems.push(val.secondaryText);
    });
    userEmails = notificationListItems.toString().split(',').join('; ');
    setValue(
      'notificationList',
      userEmails
    );
    // console.log(userEmails);
  };

  // used to manually register emailDL in react-hook-form (needed for all multiselect type dropdowns or checkboxes)
  useEffect(() => {
    // need to manually register emailDL with react-hook-form since it's multiselect
    register({ name: 'submitFlag' });
    register({ name: 'emailDL' });
    register({ name: 'notificationList' });
    register({ name: 'requestType' });
    register({ name: 'candidateStartDate' });


    user.then(res => {
      setUserInfo(res);
      if (itemIdVal === 'new') {
        setValue(
          'submittedBy',
          res.Title
        );
      }
    });

    // pnp sp ajax call to grab client list names/data
    sp.web.lists.getByTitle('Client-List')
      .items
      .select('ID', 'Title')
      .orderBy('Title', true)
      .getAll()
      .then(data => {
        setClientNames(data);
      })
      .catch(err => {
        console.log(err);
      });

    console.log('useEffect fired');

  }, []);

  // subRequestType change event handler (using onChanged component prop instead of onChange)
  const onSubRequestTypeChanged = (option) => {
    console.log(option.dataRequestType);
    setValue(
      'requestType',
      option.dataRequestType
    );

  };

  // onChange event handler for emailDL multiselect dropdown
  const handleEmailDlChange = (event, item) => {
    // setValue('emailDL', item);
    // setEmailDlKeys(item);
    if (item) {
      setEmailDlKeys(
        item.selected ? [...emailDlKeys, item.key as string] : emailDlKeys.filter(key => key !== item.key)
      );

      setValue(
        'emailDL',
        {
          "__metadata": {
            type: "Collection(Edm.String)"
          },
          results: emailDlKeys
        }
      );
    }
  };

  // function to clear date input and reset form value in state and react-hook-form
  const onClearDateInput = React.useCallback((): void => {
    setDateValue(undefined);
    setValue(
      'candidateStartDate',
      ''
    );
    console.log('onClearDateInput called');
  }, []);

  const onParseDateFromString = React.useCallback(
    (newValue: string): Date => {
      const previousValue = dateValue || new Date();
      const newValueParts = (newValue || '').trim().split('/');
      const day =
        newValueParts.length > 0 ? Math.max(1, Math.min(31, parseInt(newValueParts[0], 10))) : previousValue.getDate();
      const month =
        newValueParts.length > 1
          ? Math.max(1, Math.min(12, parseInt(newValueParts[1], 10))) - 1
          : previousValue.getMonth();
      let year = newValueParts.length > 2 ? parseInt(newValueParts[2], 10) : previousValue.getFullYear();
      if (year < 100) {
        year += previousValue.getFullYear() - (previousValue.getFullYear() % 100);
      }
      return new Date(year, month, day);
    },
    [dateValue],
  );

  // formats the appearance of the date format to the user, and manually sets the value for react-hook-form
  const onFormatDate = (date?: Date): string => {
    setValue(
      'candidateStartDate',
      !date ? '' : (('0' + (date.getMonth() + 1)).slice(-2) + '/' + ('0' + (date.getDate())).slice(-2) + '/' + date.getFullYear())
    );

    return !date ? '' : (('0' + (date.getMonth() + 1)).slice(-2) + '/' + ('0' + (date.getDate())).slice(-2) + '/' + date.getFullYear());
  };

  // TODO FINISH EVENT HANDLER LOGIC (SOME MAY BE ABLE TO BE REPLACED VIA STATE/REACT-HOOK-FORM STATE AND CONDITIONAL RENDERING)




  const onPendingReqApprovalChange = (e, item) => {
    console.log(e);
    // console.log(e.target.value);
    setPendingReqApprovalValue(e);
    if (watchPendingReqApproval === 'No') {
      onClearDateInput();
    }
  };

  // cancel event handler. Based on the state param in the address bar of the ticket, the user is relocated to a specific URL
  const onCancelClick = () => {
    switch (true) {
      case (itemStateVal === 'cim'):
        window.location.href = '/sites/teams/sa/contractimplementation/Pages/cimpage.aspx';
        break;
      case (itemStateVal === 'legal'):
        window.location.href = '/sites/teams/sa/contractimplementation/Pages/legalpage.aspx';
        break;
      case (itemStateVal === 'info' || itemStateVal === 'submit'):
        window.location.href = '/sites/teams/sa/contractimplementation/Pages/submitterpage.aspx';
        break;
      case (itemIdVal !== 'new'):
        window.location.href = '/sites/teams/sa/contractimplementation/Lists/sacc/Dashboard.aspx';
        break;
      default:
        window.location.href = '/sites/teams/sa/contractimplementation/Pages/submitterpage.aspx';
        break;
    }
  };

  // submits form data and clears validFormData object
  const onSubmit = (data) => {
    setValidationError(null);
    setValidFormData(null);

    handleSubmit(async (data) => {
      console.log(data);
      // set submitFlag to true for react-hook-form to submit to the list
      setValue(
        'submitFlag',
        true
      )
      setValidFormData(data);

      // if else to determine logic of whether it's a new item or an update to an existing item
      if (itemIdVal === 'new') {

        // add new list item
        await sp.web.lists.getByTitle('sacc')
          .items
          .add(data)
          .then(res => {
            console.log(res.data.Id);
            listItemAttachmentsComponentReference.current?.uploadAttachments(res.data.Id);
            // window.location.href = '/sites/teams/sa/contractimplementation/Pages/submitterpage.aspx';
          })
          .catch(err => {
            console.log(err);
          });

      } else {

        // update list item. Pass convert itemIdVal using parseInt
        sp.web.lists.getByTitle('sacc')
          .items
          .getById(parseInt(itemIdVal))
          .update(data)
          .then(res => {
            console.log(res);

            // listItemAttachmentsComponentReference.current.uploadAttachments(1);

            if (itemStateVal === 'submit' || itemStateVal === 'info') {
              window.location.href = '/sites/teams/sa/contractimplementation/Pages/submitterpage.aspx';
            } else if (itemStateVal === 'cim') {
              window.location.href = '/sites/teams/sa/contractimplementation/Pages/cimpage.aspx';
            } else if (itemStateVal === 'legal') {
              window.location.href = '/sites/teams/sa/contractimplementation/Pages/legalpage.aspx';
            } else {
              window.location.href = '/sites/teams/sa/contractimplementation/Lists/sacc/Dashboard.aspx';
            }
          })
          .catch(err => {
            console.log(err);
          });

      }

    }, (err) => {
      console.log('Error: ' + err);
      setValidationError(err);
    })();
  };


  // returns JSX SACC form
  return (
    <ThemeProvider theme={rhTheme}>

      <Stack horizontalAlign='center'>

        <Stack styles={stackStyles}>
          <Stack horizontalAlign='start'>
            <Text variant="xLarge" block styles={{ root: { color: rhTheme.palette.themePrimary, marginBottom: 20 } }}>
              {props.description}
            </Text>
          </Stack>
        </Stack>

        <Stack horizontal tokens={stackTokens} styles={stackStyles}>


          {/* left column */}
          <Stack {...columnProps}>

            <TextField label="Ticket #:"
              disabled
            />

            <ControlledDropdown
              required={true}
              options={clientNames.map(val => ({ key: val.Title, text: val.Title }))}
              label="Client Name:"
              control={control}
              name={nameof<Form>('clientName')}
              errors={errors}
              placeholder="Select a value"
              rules={{ required: "Please select a value" }}
            />

            <ControlledDropdown
              required={false}
              options={[
                { key: 'Strategic Account', text: 'Strategic Account' },
                { key: 'Regional Account', text: 'Regional Account' },
              ]}
              label="Account Designation:"
              control={control}
              name={nameof<Form>('acctDesignation')}
              errors={errors}
              placeholder="Select a value"
            />

            {/* react-hook-form test */}
            <ControlledTextField
              required={true}
              label="VP on Account"
              control={control}
              name={nameof<Form>('vpOnAccount')}
              errors={errors}
              rules={{ required: "This field is required" }}
            />


            <ControlledDropdown
              required={true}
              options={[
                { key: 'San Ramon', text: 'San Ramon' },
                { key: 'Columbus', text: 'Columbus' },
              ]}
              label="Client Services Team:"
              control={control}
              name={nameof<Form>('clientServiceTeam')}
              errors={errors}
              placeholder="Select a value"
              rules={{ required: "Please select a value" }}
            />


            <ControlledDropdown
              required={true}
              options={subRequestTypeOptions}
              label="Request Type:"
              control={control}
              name={nameof<Form>('subRequestType')}
              onChanged={onSubRequestTypeChanged}
              errors={errors}
              placeholder="Select a value"
              rules={{ required: "Please select a value" }}
            />

            <ControlledTextField
              required={false}
              label="Other Request Type:"
              control={control}
              name={nameof<Form>('otherRequestType')}
              errors={errors}
            // rules={{ required: "This field is required" }}
            />

            <Dropdown
              placeholder="Select Options..."
              label="Email DL:"
              selectedKeys={emailDlKeys}
              // eslint-disable-next-line react/jsx-no-bind
              onChange={handleEmailDlChange}
              multiSelect
              options={emailDlOptions}
            // styles={dropdownStyles}
            />


          </Stack>


          {/* right column */}
          <Stack {...columnProps}>

            <Link href="https://roberthalf.sharepoint.com/:x:/r/sites/teams/sa/contractimplementation/_layouts/15/Doc.aspx?sourcedoc=%7B7B054E40-62A6-44BA-85BD-F431B20CF04D%7D&file=SACC%20YES%20NO%20Grid%20-%201.16.2020.xlsx&action=default&mobileredirect=true" target='_blank' styles={{ root: { color: '#68ACE5', fontWeight: 'bold' } }}>Click to see Yes/No Grid</Link>


            <ControlledTextField
              label="Request Description:"
              control={control}
              name={nameof<Form>('requestDescription')}
              errors={errors}
              required={true}
              multiline
              rows={10}
              rules={{ required: "This field is required" }}
            />

            <ListItemAttachments
              ref={listItemAttachmentsComponentReference}
              context={props.context}
              listId="ce9fb974-42ec-4c54-bf27-babfcfffdc92"
            // itemId={0}
            />

            <ControlledTextField
              label="Requestor(Submitter):"
              control={control}
              name={nameof<Form>('submittedBy')}
              disabled
              errors={errors}
            />

            <ControlledDropdown
              required={true}
              options={[
                { key: 'VP', text: 'VP' },
                { key: 'Client', text: 'Client' },
                { key: 'Services Team', text: 'Services Team' },
                { key: 'Other', text: 'Other' },
              ]}
              label="Request Origination:"
              control={control}
              name={nameof<Form>('origination')}
              errors={errors}
              placeholder="Select a value"
              rules={{ required: "Please select a value" }}
            />

            {watchOrigination === 'Other' &&
              <ControlledTextField
                required={watchOrigination === 'Other' ? true : false} // needs to be a conditional to toggle true or false
                label="Other Origination:"
                control={control}
                name={nameof<Form>('otherOrigination')}
                errors={errors}
                rules={watchOrigination === 'Other' ? { required: "This field is required" } : null}
              />
            }

            <ControlledDropdown
              options={[
                { key: 'No', text: 'No' },
                { key: 'Yes', text: 'Yes' },
              ]}
              label="Result of Client Audit?:"
              control={control}
              name={nameof<Form>('clientAudit')}
              errors={errors}
              placeholder="Select a value"
            />


            <ControlledDropdown
              options={[
                { key: 'No', text: 'No' },
                { key: 'Yes', text: 'Yes' },
              ]}
              label="New Starts Pending Request Approval?:"
              control={control}
              name={nameof<Form>('pendingReqApproval')}
              errors={errors}
              selectedKey={pendingReqApprovalValue ? pendingReqApprovalValue.key : undefined}
              onChanged={onPendingReqApprovalChange}
              placeholder="Select a value"
            />

            {/* TODO need to assess why value gets cleared out and cannot set MM/dd/yyyy as react-hook-form value */}

            {watchPendingReqApproval === 'Yes' &&
              <>
                <DatePicker
                  componentRef={datePickerRef}
                  label="Candidate Start Date:"
                  // allowTextInput
                  ariaLabel="Select a date. Input format is day slash month slash year."
                  value={dateValue}
                  onSelectDate={setDateValue as (date?: Date) => void}
                  formatDate={onFormatDate}
                  parseDateFromString={onParseDateFromString}
                  // className={styles.control}
                  // DatePicker uses English strings by default. For localized apps, you must override this prop.
                  strings={defaultDatePickerStrings}
                />
                {/* <DefaultButton aria-label="Clear date input" onClick={onClearDateInput} text="Clear Date" styles={{ root: { width: '35%' } }} /> */}
              </>
            }

            <PeoplePicker
              context={props.context}
              titleText="Notification List"
              personSelectionLimit={3}
              groupName={""} // Leave this blank in case you want to filter from all users
              showtooltip={true}
              tooltipMessage="Start typing to see names"
              onChange={_getPeoplePickerItems}
              showHiddenInUI={false}
              principalTypes={[PrincipalType.User]}
              resolveDelay={700}
            />

            <ControlledTextField
              label="Request Details:"
              control={control}
              name={nameof<Form>('additionalInfo')}
              errors={errors}
              required={false}
              multiline
              rows={10}
            />

            <ControlledTextField
              label="Request Details Summary (Read Only):"
              control={control}
              name={nameof<Form>('additionalInfoSummary')}
              errors={errors}
              required={false}
              multiline
              readOnly
              description='Read Only'
              rows={10}
            />

          </Stack>

        </Stack>

        <Stack styles={stackStyles}>
          <Stack horizontalAlign='start'>
            <Text variant="xLarge" block styles={{ root: { color: rhTheme.palette.themePrimary, marginBottom: 20 } }}>
              Reviewer Information
            </Text>
          </Stack>
        </Stack>

        <Stack horizontal tokens={stackTokens} styles={{ root: { width: 990, paddingTop: 20 } }}>


          <Stack {...columnProps}>

            <ControlledDropdown
              options={[
                { key: 'New', text: 'New' },
                { key: 'Assigned to RH Connect Team', text: 'Assigned to RH Connect Team' },
                { key: 'Assigned to SA Setup Team', text: 'Assigned to SA Setup Team' },
                { key: 'In Progress', text: 'In Progress' },
                { key: 'In CIM Review', text: 'In CIM Review' },
                { key: 'In Legal Review', text: 'In Legal Review' },
                { key: 'Completed', text: 'Completed' },
                { key: 'Closed-RHCS', text: 'Closed-RHCS' },
                { key: 'On Hold', text: 'On Hold' },
                { key: 'Denied', text: 'Denied' },
                { key: 'Cancelled', text: 'Cancelled' },
                { key: 'Pending Additional Information', text: 'Pending Additional Information' },
              ]}
              label="Ticket Status:"
              control={control}
              required={true}
              name={nameof<Form>('status')}
              errors={errors}
              placeholder="Select a value"
              rules={{ required: "This field is required" }}
            />

            <ControlledDropdown
              options={[
                { key: 'Andrea Mahoney', text: 'Andrea Mahoney' },
                { key: 'Anita Sandher', text: 'Anita Sandher' },
                { key: 'Blake Buttram', text: 'Blake Buttram' },
                { key: 'Carolyn Barbagelata', text: 'Carolyn Barbagelata' },
                { key: 'Jennifer Camp', text: 'Jennifer Camp' },
                { key: 'Jennifer Dutra', text: 'Jennifer Dutra' },
                { key: 'Myrisha Howard', text: 'Myrisha Howard' },
                { key: 'Raul Guillarte', text: 'Raul Guillarte' },
                { key: 'Sarah Simmons', text: 'Sarah Simmons' },
                { key: 'Shawna Bailey', text: 'Shawna Bailey' },
                { key: 'Snneha Naarang', text: 'Snneha Naarang' },
                { key: 'Stephen Hoshida', text: 'Stephen Hoshida' },
              ]}
              label="CIM Reviewer:"
              control={control}
              name={nameof<Form>('cimReviewer')}
              errors={errors}
              placeholder="Select a value"
            />

            <ControlledDropdown
              options={[
                { key: 'Alison King', text: 'Alison King' },
                { key: 'Amy Ogdie', text: 'Amy Ogdie' },
                { key: 'Annie Betinol', text: 'Annie Betinol' },
                { key: 'Asalya Winger', text: 'Asalya Winger' },
                { key: 'Azima Subedar', text: 'Azima Subedar' },
                { key: 'Cameron Fortner', text: 'Cameron Fortner' },
                { key: 'Cassandra Bauccio', text: 'Cassandra Bauccio' },
                { key: 'Ruba Forno', text: 'Ruba Forno' },
                { key: 'Ian Henri', text: 'Ian Henri' },
                { key: 'Jason Castello', text: 'Jason Castello' },
                { key: 'Juliet Jonas', text: 'Juliet Jonas' },
                { key: 'Jennifer Daniele', text: 'Jennifer Daniele' },
                { key: 'Joe Aguilar', text: 'Joe Aguilar' },
                { key: 'Katina Sharp', text: 'Katina Sharp' },
                { key: 'Maisie Cole', text: 'Maisie Cole' },
                { key: 'Michael Balistreri', text: 'Michael Balistreri' },
                { key: 'Nick Young', text: 'Nick Young' },
                { key: 'Shaul Serban', text: 'Shaul Serban' },
              ]}
              label="Legal Reviewer:"
              control={control}
              name={nameof<Form>('legalReviewer')}
              errors={errors}
              placeholder="Select a value"
            />

            <ControlledDropdown
              options={[
                { key: 'Bela Cavaleiro', text: 'Bela Cavaleiro' },
                { key: 'Chloe Sheard', text: 'Chloe Sheard' },
                { key: 'Christian Aiken', text: 'Christian Aiken' },
                { key: 'Curtis Wilson', text: 'Curtis Wilson' },
                { key: 'Monica Irwin', text: 'Monica Irwin' },
                { key: 'Nancie Malucchi', text: 'Nancie Malucchi' },
                { key: 'Paula Garcia', text: 'Paula Garcia' },
                { key: 'Shannon Ramirez', text: 'Shannon Ramirez' },
                { key: 'Sylvia Nix', text: 'Sylvia Nix' },
              ]}
              label="SA Reviewer:"
              control={control}
              name={nameof<Form>('saReviewer')}
              errors={errors}
              placeholder="Select a value"
            />

          </Stack>


          <Stack {...columnProps}>


            <ControlledDropdown
              options={[
                { key: 'Approved', text: 'Approved' },
                { key: 'Not Approved', text: 'Not Approved' },
                { key: 'Need Further Information', text: 'Need Further Information' },
                { key: 'Need Further Information - SACC', text: 'Need Further Information - SACC' },
              ]}
              label="Legal Reviewer Response:"
              control={control}
              name={nameof<Form>('legalReviewerResponse')}
              errors={errors}
              placeholder="Select a value"
            />

            <ControlledDropdown
              options={[
                { key: 'In Progress', text: 'In Progress' },
                { key: 'QA Ready', text: 'QA Ready' },
                { key: 'Need Further Information', text: 'Need Further Information' },
              ]}
              label="CIM Reviewer Response:"
              control={control}
              name={nameof<Form>('cimReviewerResponse')}
              errors={errors}
              placeholder="Select a value"
            />

            <ControlledTextField
              label="Internal Notes:"
              control={control}
              name={nameof<Form>('internalNotes')}
              errors={errors}
              required={false}
              multiline
              rows={10}
            />

            <ControlledTextField
              label="Internal Notes Summary (Read Only):"
              control={control}
              name={nameof<Form>('internalNoteSummary')}
              errors={errors}
              required={false}
              multiline
              readOnly
              rows={10}
            />

          </Stack>

        </Stack>

        <Stack horizontal horizontalAlign='center' tokens={stackTokens} styles={{ root: { width: 990, paddingTop: 20 } }}>

          <Stack styles={{ root: { width: 150 } }}>

            <PrimaryButton onClick={onSubmit} text="Submit" />
          </Stack>

          <Stack styles={{ root: { width: 150 } }}>

            <DefaultButton onClick={onCancelClick} text="Cancel" />

          </Stack>


        </Stack>

        {validationError && (
          <>
            <div>Form validation errors:</div>
            <div><pre>{JSON.stringify(validationError, null, 2)}</pre></div>
          </>
        )}
        {validFormData && (
          <>
            <div>Form passed all validations</div>
            <div><pre>{JSON.stringify(validFormData, null, 2)}</pre></div>
          </>
        )}

      </Stack>
    </ThemeProvider>
  );

};

export default Sacc;
