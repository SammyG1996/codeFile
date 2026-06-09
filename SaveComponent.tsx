import * as React from "react";
import { Button, Spinner } from "@fluentui/react-components";
import { DynamicFormContext } from "@spfx-monorepo/shared-library/dist/cjs/components/DynamicFormContext";
import postSPRestAPI, {
  ReturnDataProps,
} from "@spfx-monorepo/shared-library/dist/cjs/Utils/postSPRestAPI";
import { FormCustomizerContext } from "@microsoft/sp-listview-extensibility";
import { StepId } from "../flowScaffold/types";
import { processMap } from "../flowScaffold/processMap";
import type { UserProps } from "@spfx-monorepo/shared-library/dist/cjs/Utils/types";
import { createFlowEngine } from "../flowScaffold/engine";
import type {
  RequestTracker,
  InternalFieldNames,
  StatusChoices,
  RequestHistoryEntry,
} from "../flowScaffold/types";
import { decisionExecuter } from "../flowScaffold/deciders";
import {
  buildEmail,
  EmailPayload,
  EmailRouterContext,
  FlowBody,
  FlowResult,
  sendEmail,
} from "../flowScaffold/email";
import { evaluateFieldRules } from "@spfx-monorepo/shared-library/dist/cjs/Utils/formRulesEngine";

interface ButtonProps {
  OnSubmit: (data: boolean) => void; // Parent submit toggle
  submitting: boolean; // Disable while submitting
  formContext: FormCustomizerContext; // SPFx form context
  selectedType: string;
  onSave: () => void;
  onClose: () => void;
}

export default function SaveComponent(props: ButtonProps): JSX.Element {
  /* ---------- Context + UI state ---------- */
  const ctx = DynamicFormContext();
  const [isHidden, setIsHidden] = React.useState<boolean>(false);
  const [isDisabled, setIsDisabled] = React.useState<boolean>(false);
  const [spinnerHidden, setSpinnerHidden] = React.useState<boolean>(true);
  const [spinnerLabel, setSpinnerLabel] = React.useState<string>("");
  const btnId = "btnSubmit";

  type pplPickerStorage = {
    email: string;
    fullName: string;
  };

  /* ---------- Form modes ---------- */
  const FORM_MODE_DISPLAY = 4;
  const FORM_MODE_EDIT = 6;
  const FORM_MODE_NEW = 8;

  /* ---------- Small UI helpers ---------- */
  const showSpinner = (label: string) => {
    setSpinnerLabel(label);
    setSpinnerHidden(false);
  };

  const hideSpinner = () => {
    setSpinnerLabel("");
    setSpinnerHidden(true);
  };

  const endSubmitUI = () => {
    props.OnSubmit(false);
    hideSpinner();
  };

  const endSaveSuccessUI = () => {
    props.OnSubmit(false);
    hideSpinner();
    const currentUser: UserProps = ctx.curUserInfo;
    const groups = currentUser?.SPGroups ?? [];
    const internalUserCheck = groups.some(
      (g: any) => g?.Title === "Knowledge Management Internal Users"
    );
    const ownerUserCheck = groups.some(
      (g: any) => g?.Title === "Request Form Owners"
    );

    const shouldRedirectToThankYou =
      ctx.FormMode === FORM_MODE_NEW ||
      (ctx.FormMode === FORM_MODE_EDIT &&
        !internalUserCheck &&
        !ownerUserCheck);

    if (shouldRedirectToThankYou) {
      window.location.href =
        "https://amerihealthcaritas.sharepoint.com/sites/eokm/SitePages/Thank-you.aspx?env=Embedded";
    } else {
      // ✅ Let SPFx do its normal redirect to the List
      props.onSave();
    }
  };

  //For EndPoint for Attachments
  const buildItemEndPoint = (itemID: number) => {
    return `${props.formContext.pageContext.site.serverRelativeUrl}/_api/web/lists/GetByTitle('${props.formContext.list.title}')/items(${itemID})`;
  };

  //Configuration to PICK POST vs PATCH
  const getRequestConfig = () => {
    // New item
    if (ctx.FormMode === FORM_MODE_NEW) {
      return {
        // Return config
        url: `${props.formContext.pageContext.site.serverRelativeUrl}/_api/web/lists/GetByTitle('${props.formContext.list.title}')/items`, // Create
        method: "POST" as const, // POST
      };
    }
    // Existing ID
    const id = props.formContext.item?.ID as number;
    return {
      // Return config
      url: `${props.formContext.pageContext.site.serverRelativeUrl}/_api/web/lists/GetByTitle('${props.formContext.list.title}')/items(${id})`, // Update
      method: "PATCH" as const, // MERGE via PATCH
    };
  };

  //Function for form validations
  const validateForm = async (): Promise<boolean> => {
    /*
     * 1 Check for required fields have values are not empty
     * 2 Check for any outstanding error messages
     */
    const res: any = ctx.GlobalReturnData(); // eslint-disable-line @typescript-eslint/no-explicit-any
    const errorItems = res?.errorItems ?? {}; // For any outstanding errors across all input fields
    const listData = res?.listData ?? {}; // For any input elements change values - Only Elements that have been touched or updated
    const frmData = res?.frmData ?? {}; //Available only for EditForm/Existing values of all input fields
    const requiredElements: string[] = res?.requiredItems ?? []; // List of Required Input Elements

    // Treat null/undefined/[]/{results:[]} as empty
    const isEmpty = (v: any): boolean => {
      if (v === null || v === undefined) return true;
      if (Array.isArray(v)) return v.length === 0;
      if (typeof v === "object" && Array.isArray(v.results))
        return v.results.length === 0;
      return false;
    };

    for (const field of requiredElements) {
      const value =
        ctx.FormMode === FORM_MODE_EDIT ? frmData[field] : listData[field]; //Source by mode
      if (isEmpty(value)) {
        //Required field is missing value
        alert("Please fill in all required fields!"); //Notify
        return false; //Fail
      }
    }

    if (Object.keys(errorItems).length > 0) {
      // Any errors?
      alert("Please review highlighted fields and try again!"); // Notify
      return false; //Fail
    }
    return true; //Pass
  };

  // Function to upload files (SEQUENTIAL - no save conflict)
  const uploadAttachments = async (attachments: any[], itemId: any) => {
    // eslint-disable-line @typescript-eslint/no-explicit-any
    if (!attachments || attachments.length === 0) return [];

    const itemUrl = buildItemEndPoint(itemId);

    const headers: HeadersInit = {
      Accept: "application/json",
      "Content-Type": "application/octet-stream",
    };

    const results: any[] = [];

    for (const att of attachments) {
      const fileUrl = `${itemUrl}/AttachmentFiles/add(FileName='${encodeURIComponent(
        att.name
      )}')`;

      const apiCall = {
        uri: fileUrl,
        body: att.content,
        context: props.formContext,
      };

      try {
        const res = await postSPRestAPI(apiCall, headers);
        results.push({ status: "fulfilled", value: res });
      } catch (err) {
        results.push({ status: "rejected", reason: err });
      }

      // Optional but recommended: small delay to avoid item lock timing issues
      await new Promise((resolve) => setTimeout(resolve, 200));
    }

    return results;
  };

  //Strip attachment when there are no attachments
  const stripAttachments = (data: Record<string, any>) => {
    // Remove attachments
    const copy = { ...data }; // Clone
    delete copy.attachments; // Drop attachments key
    return copy; // Return cleaned object
  };

  //Save item to SP
  const saveItem = async (): Promise<ReturnDataProps> => {
    const { url, method } = getRequestConfig(); //Get URL/method
    const raw = ctx.ListSubData ?? {}; //Source data
    const deepCopy = JSON.parse(JSON.stringify(raw)); // Deep Copy
    const dataToSend = stripAttachments(deepCopy);
    const headers: HeadersInit =
      method === "PATCH" //Headers by method
        ? {
            // PATCH (MERGE)
            Accept: "application/json;odata.metadata=full", // Response type
            "Content-Type": "application/json;odata.metadata=full", // Body type
            "X-HTTP-Method": "MERGE", // MERGE semantics
            "IF-MATCH": "*", // Ignore version
          }
        : {
            // POST
            Accept: "application/json;odata.metadata=full", // Response type
            "Content-Type": "application/json;odata.metadata=full", // Body type
          };
    const apiCall = {
      uri: url,
      body: JSON.stringify(dataToSend),
      context: props.formContext,
    };
    const result = await postSPRestAPI(apiCall, headers); // Execute call

    if (result.status !== 201 && result.status !== 204) {
      // Accept only success
      throw result; // Bubble up error
    }
    return result; // Return success result
  };

  //Function to Normalize RequestTracker so that old format is updated to new format

  // --- OLD SHAPE (minimal) ---
  type OldRequestTrackerItem = {
    RqstrDetails: { RqstrName: string; RqstrEmail: string; RqstrId: number };
    Status: Record<string, string>;
    PrevStatus?: string;
    TWDetails?: { TWName: string; TWEmail: string };
    TLDetails?: { TLName: string; TLEmail: string };
  };

  // --- TYPE GUARDS ---
  const isOld = (x: any): x is OldRequestTrackerItem =>
    !!x?.RqstrDetails &&
    typeof x?.Status === "object" &&
    !Array.isArray(x?.statusDetails);

  const isNew = (x: any): x is RequestTracker =>
    !!x?.requestorDetails && Array.isArray(x?.statusDetails);

  // --- REQUESTTRACKER NORMALIZER (old -> new only; new passes through) ---
  /** This will update the keys to new names for any old requests that were submitted before this */
  const normalizeRequestTracker = (data: any) => {
    // If already normalized
    if (data?.requestor && Array.isArray(data?.history)) return data;

    const o = Array.isArray(data) ? data[0] : data;

    const toIsoUTC = (s: string) =>
      new Date(
        new Date(s.replace(";", " ")).toLocaleString("en-US", {
          timeZone: "America/New_York",
        })
      ).toISOString();

    const stepId = (status: string) =>
      ({
        Open: "P100",
        "Assigned to TW": "P500",
        "Route to Team Lead": "P400",
        "Route to Submitter": "P300",
        Cancelled: "P600",
      } as Record<string, string>)[status] ?? "UNKNOWN";
    const entries = Object.entries(o.Status as Record<string, string>) as [
      string,
      string
    ][];

    const history = entries
      .sort(
        ([, a], [, b]) =>
          Date.parse(a.replace(";", " ")) - Date.parse(b.replace(";", " "))
      )
      .map(([status, ts]) => ({
        stepId: stepId(status),
        status,
        timestamp: toIsoUTC(ts),
        by: o.PrevStatusModBy || o.RqstrDetails.RqstrName,
        ...(status === "Assigned to TW" &&
          o.TWDetails && {
            technicalWriterName: o.TWDetails.TWName,
            technicalWriterEmail: o.TWDetails.TWEmail,
          }),
        ...(status === "Route to Submitter" &&
          o.RouterDetails && {
            routerName: o.RouterDetails.RName,
            routerEmail: o.RouterDetails.REmail,
          }),
        ...(status === "Route to Team Lead" &&
          o.TLDetails && {
            teamLeadName: o.TLDetails.TLName,
            teamLeadEmail: o.TLDetails.TLEmail,
          }),
      }));

    return {
      requestor: {
        name: o.RqstrDetails.RqstrName,
        email: o.RqstrDetails.RqstrEmail,
        spId: o.RqstrDetails.RqstrId,
      },
      history,
    };
  };

  //If internal users accidentally changes Assigned To or Assigned To Team Lead values when it is not appropriate, then it will restore to previous value
  const enforceAssignmentLocks = (
    nextStepId: StepId,
    flowTracker: RequestTracker
  ) => {
    // Steps where change IS allowed
    const twAllowedSteps: StepId[] = ["P500", "P1100", "P1600"]; // Assign / Reassign TW
    const tlAllowedSteps: StepId[] = ["P400", "P900", "P1400"]; // Route to Team Lead

    // ---------------- Assigned To (TW) ----------------
    if (
      !twAllowedSteps.includes(nextStepId) &&
      "Assigned_x0020_ToId" in ctx.ListSubData
    ) {
      const lastTW = [...flowTracker.history]
        .reverse()
        .find(
          (h) => h.status === "Assigned to TW" || h.status === "Reassign to TW"
        );
      const pplPickerDetails = getPeoplePickerInfo(
        String(ctx.ListSubData["Assigned_x0020_ToId"]),
        "Assigned_x0020_To"
      );

      if (lastTW?.technicalWriterEmail !== pplPickerDetails?.email) {
        //IMPORTANT: remove user change entirely
        delete ctx.ListSubData["Assigned_x0020_ToId"];
        alert(
          "Your changes to Assigned To field will not be saved unless Status is Assigned To TW or Reassign to TW!"
        );
      }
    }

    // ---------------- Team Lead ----------------
    if (
      !tlAllowedSteps.includes(nextStepId) &&
      "Assign_x0020_to_x0020_team_x0020Id" in ctx.ListSubData
    ) {
      const lastTL = [...flowTracker.history]
        .reverse()
        .find((h) => h.status === "Route to Team Lead");
      const pplPickerDetails = getPeoplePickerInfo(
        String(ctx.ListSubData["Assign_x0020_to_x0020_team_x0020Id"]),
        "Assigned_x0020_To"
      );

      if (lastTL?.teamLeadEmail !== pplPickerDetails?.email) {
        delete ctx.ListSubData["Assign_x0020_to_x0020_team_x0020Id"];
        alert(
          "Your changes to Assign to team lead field will not be saved unless Status is Route to Team Lead!"
        );
      }
    }
  };

  //Get Local Storage for People Picker Ids
  const getPeoplePickerInfo = (
    pickerId: string,
    fieldName: InternalFieldNames
  ): pplPickerStorage | null => {
    const localStoragevar = `${props.formContext.pageContext.web.title}.peoplePickerIDs.${fieldName}`;
    const storedRaw = localStorage.getItem(localStoragevar) ?? "[]";
    const storedArr = JSON.parse(storedRaw) as any[]; // eslint-disable-line @typescript-eslint/no-explicit-any
    const checkSPUserEmail =
      (pickerId &&
        (storedArr.find((x) => x?.EntityData?.SPUserID === pickerId)
          ?.Email as string)) ??
      null;
    const checkSPUserName =
      (pickerId &&
        (storedArr.find((x) => x?.EntityData?.SPUserID === pickerId)
          ?.DisplayText as string)) ??
      null;

    if (checkSPUserEmail !== null || checkSPUserName !== null) {
      const SPUserFullName = `${checkSPUserName.split(", ")[1]} ${
        checkSPUserName.split(", ")[0]
      }`;
      return { email: checkSPUserEmail, fullName: SPUserFullName };
    } else {
      return null;
    }
  };

  type StepResult = { ok: boolean; value?: EmailRouterContext };
  //Step handlers executes all Flow steps using engine and updates necessary fields
  const stepHandler = async (): Promise<StepResult> => {
    //Create Flow Engine instance
    const engine = createFlowEngine(processMap);
    const currentUser: UserProps = ctx.curUserInfo; //Current user Email and name

    // Retrieve the RequestTracker field value from existing list item (Will be empty for NewForm)
    const rawTracker: RequestTracker =
      ctx.FormMode !== FORM_MODE_NEW
        ? JSON.parse(ctx.FormData?.RequestTracker)
        : {};
    //Normalize the RequestTracker value so that it is aligned to more user friendly property keynames - this will handle Old requests

    const flowTracker: RequestTracker =
      rawTracker && Object.keys(rawTracker).length
        ? normalizeRequestTracker(rawTracker)
        : {};

    flowTracker.history = flowTracker.history ?? [];
    const lastEntry = flowTracker.history.at(-1);

    let currentStatusval: StatusChoices = ctx.FormData?.Status ?? null; //Status field value
    // If was routed and the right person is acting now, set Status as "Returned from route"
    if (
      lastEntry?.status === "Route to Submitter" &&
      currentUser.userEmail === flowTracker.requestor.email
    ) {
      currentStatusval = "Returned from Route";
    }
    if (
      lastEntry?.status === "Route to Team Lead" &&
      currentUser.userEmail === lastEntry.teamLeadEmail
    ) {
      currentStatusval = "Returned from Route";
    }

    //For new form, currentstepid is always start step id P100
    const currentStepId: StepId =
      ctx.FormMode === FORM_MODE_NEW
        ? processMap.startStepId
        : (lastEntry?.stepId as StepId);

    let nextStepId: StepId = processMap.startStepId; //default for new form
    if (ctx.FormMode !== FORM_MODE_NEW) {
      /** For Edit Form check if the next step is a decision step */
      // const nextCandidate = engine.next(currentStepId);
      const nextCandidate = engine.next(currentStepId, currentStatusval, {
        decide: decisionExecuter,
      });
      if (!nextCandidate) return { ok: true }; //This means that we have reached to the end of the Process map so just return back

      const nextStep = engine.getStep(nextCandidate);
      const isDecision = nextStep.shapeType.trim().toLowerCase() === "decision";
      nextStepId = isDecision
        ? engine.next(nextCandidate, currentStatusval, {
            decide: decisionExecuter,
          }) ?? nextCandidate
        : nextCandidate;
      //Below will make sure that for EditForms the contentType is added
      ctx.ListSubData = {
        ...ctx.ListSubData,
        ContentTypeId: ctx.FormData.ContentTypeId,
      };
    }

    if (lastEntry?.stepId === nextStepId) return { ok: true }; //Prevents from re-running switch for same steps again and again

    // -------------------------
    // helper functions
    // -------------------------
    const saveTracker = () => {
      ctx.ListSubData = {
        ...ctx.ListSubData,
        RequestTracker: JSON.stringify(flowTracker),
      };
    };
    const addHistory = (entry: RequestHistoryEntry) => {
      flowTracker.history.push(entry);
      saveTracker();
    };
    const currenDateTimeUser = (): string => {
      const curDate = new Date()
        .toLocaleString("en-US", { timeZone: "America/New_York" })
        .split(", ")[0];
      const curTime = new Date()
        .toLocaleString("en-US", { timeZone: "America/New_York" })
        .split(", ")[1];
      return `${curDate};${curTime}; ${currentUser.userFullName}`;
    };

    const commentsHandler = (): boolean => {
      //Use ctx.FormData instead of ctx.ListSubData because Comments History element will never be in ListSubData as it will never be touched
      const data = ctx.FormData as Record<string, any>;
      const prevCommentHistory: string | null =
        "Comment_x0020_History" in data ? data["Comment_x0020_History"] : null;
      const comments: string | null =
        "Internal_x0020__x0020_Comments" in data
          ? data["Internal_x0020__x0020_Comments"]
          : null;
      //Comments is required for routes
      if (comments === null) {
        return false;
      } else {
        const buildComments =
          `${currenDateTimeUser()} - ${comments}` +
          (prevCommentHistory ? `\n${prevCommentHistory}` : "");
        //Update ListSubData with Comments as empty and t Comments History to new value
        ctx.ListSubData = {
          ...ctx.ListSubData,
          Internal_x0020__x0020_Comments: null,
          Comment_x0020_History: buildComments,
        };
        return true;
      }
    };

    //Switch Statement to handle the cases for each step id
    enforceAssignmentLocks(nextStepId, flowTracker);
    switch (nextStepId) {
      //Status = Open ---- NewForm
      case "P100": {
        const tracker: RequestTracker = {
          requestor: {
            name: currentUser.userFullName,
            email: currentUser.userEmail,
            spId: currentUser.SPID,
          },
          history: [
            {
              stepId: "P100",
              status: "Open",
              timestamp: new Date().toISOString(),
              modifiedby: currentUser.userFullName,
            },
          ],
        };

        //If Resolution is selected accidentally, it will be removed
        delete (ctx.ListSubData as any)["Resolution"];

        ctx.ListSubData = {
          ...ctx.ListSubData,
          RequestTracker: JSON.stringify(tracker),
          Status: "Open", // Set default Status field values
          Issue_x0020_Type: props.selectedType, // Set default issue type to the content name
          Title: currentUser.userFullName,
          Department: currentUser.Dept,
        };

        // -------------------
        // Build Email context
        // -------------------
        const emails: EmailRouterContext = {
          status: "Open",
          requesterName: currentUser.userFullName,
          requesterEmail: currentUser.userEmail,
          requestTypeText: props.selectedType,
          formContext: props.formContext,
        };
        return { ok: true, value: emails };
      }
      //Status = Route to Submitter
      case "P300":
      case "P800":
      case "P1500": {
        const comments = ctx.FormData["Internal_x0020__x0020_Comments"];
        const checkComments: boolean = commentsHandler();
        if (checkComments === false) {
          alert("Please enter brief notes in Comments field!");
          return { ok: false };
        }

        addHistory({
          stepId: nextStepId,
          status: "Route to Submitter",
          timestamp: new Date().toISOString(),
          modifiedby: currentUser.userFullName,
          routerEmail: currentUser.userEmail,
          routerName: currentUser.userFullName,
        });
        // -------------------
        // Build Email context
        // -------------------
        const data = ctx.FormData as Record<string, any>; // Using FormData instead of ListSubData as Issue Type will never be touched
        const emails: EmailRouterContext = {
          status: "Route to Submitter",
          routerName: currentUser.userFullName,
          routerEmail: currentUser.userEmail,
          requesterName: flowTracker.requestor.name,
          requesterEmail: flowTracker.requestor.email,
          requestTypeText: data["Issue_x0020_Type"],
          formContext: props.formContext,
          intrnlCmmTxt4Eml: comments,
        };
        //If Resolution is selected accidentally, it will be removed
        delete (ctx.ListSubData as any)["Resolution"];
        return { ok: true, value: emails };
      }
      //Status = Route to Team Lead
      case "P400":
      case "P900":
      case "P1400": {
        const checkComments: boolean = commentsHandler();
        if (checkComments === false) {
          alert("Please enter brief notes in Comments field!");
          return { ok: false };
        }

        const field: InternalFieldNames = "Assign_x0020_to_x0020_team_x0020"; //Internal field name for Assigned to Team Lead
        //Check if Assigned To Team Lead field actual exists in ListSubdata
        if (`${field}Id` in ctx.ListSubData) {
          const data = ctx.ListSubData as Record<string, any>;
          const pplPickerValue = data[`${field}Id`]; //Retrieve Assigned to Team Lead value from ListSubData based on internal field name
          if (pplPickerValue !== null) {
            const pplPickerDetails: pplPickerStorage | null =
              getPeoplePickerInfo(String(pplPickerValue), field);
            //Check if Local storage returned team Lead values
            if (pplPickerDetails === null) {
              alert("Assigned to Team Lead cannot be empty!");
              return { ok: false };
            } else {
              const comments = ctx.FormData["Internal_x0020__x0020_Comments"];
              const checkComments: boolean = commentsHandler();
              if (checkComments === false) {
                alert("Please enter brief notes in Comments field!");
                return { ok: false };
              }

              addHistory({
                stepId: nextStepId,
                status: "Route to Team Lead",
                timestamp: new Date().toISOString(),
                teamLeadName: pplPickerDetails.fullName,
                teamLeadEmail: pplPickerDetails.email,
                modifiedby: currentUser.userFullName,
                routerName: currentUser.userFullName,
                routerEmail: currentUser.userEmail,
              });
              // -------------------
              // Build Email context
              // -------------------
              const formDataValue = ctx.FormData as Record<string, any>; // Using FormData instead of ListSubData as Issue Type will never be touched
              const emails: EmailRouterContext = {
                status: "Route to Team Lead",
                requesterName: flowTracker.requestor.name,
                routerName: currentUser.userFullName,
                routerEmail: currentUser.userEmail,
                requestTypeText: formDataValue["Issue_x0020_Type"],
                teamLeadName: pplPickerDetails.fullName,
                teamLeadEmail: pplPickerDetails.email,
                formContext: props.formContext,
                intrnlCmmTxt4Eml: comments,
              };
              //If Resolution is selected accidentally, it will be removed
              delete (ctx.ListSubData as any)["Resolution"];
              return { ok: true, value: emails };
            }
          } else {
            alert("Assigned to Team Lead cannot be empty!");
            return { ok: false };
          }
        } else {
          //If ListSubData doesn't contain Assigned to Team Lead then the field is empty.
          alert("Assigned to Team Lead cannot be empty!");
          return { ok: false };
        }
      }
      //Assigned to TW
      case "P500": {
        const field: InternalFieldNames = "Assigned_x0020_To";

        //Check if Assigned To field actual exists in ListSubdata
        if (`${field}Id` in ctx.ListSubData) {
          const data = ctx.ListSubData as Record<string, any>;
          const pplPickerValue = data[`${field}Id`]; //Retrieve Assigned to value from ListSubData based on internal field name
          if (pplPickerValue !== null) {
            const pplPickerDetails: pplPickerStorage | null =
              getPeoplePickerInfo(String(pplPickerValue), field);
            if (pplPickerDetails === null) {
              alert(
                "Assigned To cannot be empty, please enter technical writer!"
              );
              return { ok: false };
            } else {
              commentsHandler();
              addHistory({
                stepId: "P500",
                status: "Assigned to TW",
                timestamp: new Date().toISOString(),
                modifiedby: currentUser.userFullName,
                technicalWriterName: pplPickerDetails.fullName,
                technicalWriterEmail: pplPickerDetails.email,
              });
              // -------------------
              // Build Email context
              // -------------------
              const formDataValue = ctx.FormData as Record<string, any>; // Using FormData instead of ListSubData as Issue Type will never be touched
              const emails: EmailRouterContext = {
                status: "Assigned to TW",
                requesterName: flowTracker.requestor.name,
                routerName: currentUser.userFullName,
                routerEmail: currentUser.userEmail,
                requestTypeText: formDataValue["Issue_x0020_Type"],
                twName: pplPickerDetails.fullName,
                twEmail: pplPickerDetails.email,
                formContext: props.formContext,
              };
              //If Resolution is selected accidentally, it will be removed
              delete (ctx.ListSubData as any)["Resolution"];
              return { ok: true, value: emails };
            }
          } else {
            alert("Assigned to cannot be empty!");
            return { ok: false };
          }
        } else {
          alert("Assigned to cannot be empty!");
          return { ok: false };
        }
      }
      //Reassigned to TW
      case "P1100":
      case "P1600": {
        const field: InternalFieldNames = "Assigned_x0020_To";
        //Check if Assigned To field actual exists in ListSubdata
        if (`${field}Id` in ctx.ListSubData) {
          const pplPickerElm = Object.entries(ctx.ListSubData).filter(
            ([key]) => key === `${field}Id`
          )[0];
          const pplPickerid: number =
            typeof pplPickerElm[1] === "number"
              ? pplPickerElm[1]
              : Number(pplPickerElm[1]); // 0 position is always key and 1 is always the value of the key
          //Check if Assigned To field is selected but the value in the field is not empty
          if (pplPickerid !== null && pplPickerid !== 0) {
            const pplPickerDetails: pplPickerStorage | null =
              getPeoplePickerInfo(String(pplPickerid), field);
            if (pplPickerDetails === null) {
              alert("Assigned to cannot be empty!");
              return { ok: false };
            } else {
              /**
               * 1. Verify if there were any previous Reassign to TW entries, if there are more than one then pull the most recent one using timestamp
               * 2. Compare the recent Reassign to TW entry email address to the Assigned To field, if matches then throw error
               * 3. If there are no entries for previous reassign then pull the Assigned To TW entry
               * 4. Compare the Assigned to TW email to Assigned To Field, if matches then throw error
               */
              const assignTechnicalWriter = flowTracker.history.filter(
                (v) => v.status === "Assigned to TW"
              )[0];
              const reassignTechnicalWriter = flowTracker.history.filter(
                (v) => v.status === "Reassign to TW"
              );
              //Check if there was previous Reassign to TW first
              if (reassignTechnicalWriter.length > 0) {
                //if there are multiple reassign to tw then find the recent one
                const recentReassignEntry = reassignTechnicalWriter.sort(
                  (a, b) => {
                    const latest = new Date(b.timestamp);
                    const oldest = new Date(a.timestamp);
                    return latest.getTime() - oldest.getTime();
                  }
                );
                const recentReassignTWEmail =
                  recentReassignEntry[0].technicalWriterEmail;
                //Check if Recent Reassign to TW email is not same for
                if (recentReassignTWEmail === pplPickerDetails.email) {
                  alert(
                    "You cannot reassign the request to previously assigned Technical Writer, please select different!"
                  );
                  return { ok: false };
                }
              } else {
                //NO previous reassign to TW then it means there is only Assigned to TW
                if (
                  assignTechnicalWriter.technicalWriterEmail ===
                  pplPickerDetails.email
                ) {
                  alert(
                    "You cannot reassign the request to previously assigned Technical Writer, please select different!"
                  );
                  return { ok: false };
                }
              }
              commentsHandler();
              addHistory({
                stepId: nextStepId,
                status: "Reassign to TW",
                timestamp: new Date().toISOString(),
                modifiedby: currentUser.userFullName,
                technicalWriterName: pplPickerDetails.fullName,
                technicalWriterEmail: pplPickerDetails.email,
              });
              // -------------------
              // Build Email context
              // -------------------
              const formDataValue = ctx.FormData as Record<string, any>; // Using FormData instead of ListSubData as Issue Type will never be touched
              const emails: EmailRouterContext = {
                status: "Reassign to TW",
                requesterName: flowTracker.requestor.name,
                routerName: currentUser.userFullName,
                routerEmail: currentUser.userEmail,
                requestTypeText: formDataValue["Issue_x0020_Type"],
                twName: pplPickerDetails.fullName,
                twEmail: pplPickerDetails.email,
                formContext: props.formContext,
              };
              //If Resolution is selected accidentally, it will be removed
              delete (ctx.ListSubData as any)["Resolution"];
              return { ok: true, value: emails };
            }
          } else {
            alert("Assigned to cannot be empty!");
            return { ok: false };
          }
        } else {
          alert("Assigned to cannot be empty!");
          return { ok: false };
        }
      }
      //Status = In Progress
      case "P1000": {
        commentsHandler();
        addHistory({
          stepId: "P1000",
          status: "In Progress",
          timestamp: new Date().toISOString(),
          modifiedby: currentUser.userFullName,
        });
        //If Resolution is selected accidentally, it will be removed
        delete (ctx.ListSubData as any)["Resolution"];
        return { ok: true };
      }
      //Status = Completed || Cancelled
      case "P1700":
      case "P1800": {
        const data = ctx.ListSubData as Record<string, any>;
        const resolution = data["Resolution"];
        if (
          resolution === undefined ||
          resolution.length === 0 ||
          resolution === null
        ) {
          alert("Please select Resolution!");
          return { ok: false };
        }
        commentsHandler();
        addHistory({
          stepId: nextStepId,
          status: currentStatusval === "Completed" ? "Completed" : "Cancelled",
          timestamp: new Date().toISOString(),
          modifiedby: currentUser.userFullName,
        });
        // -------------------
        // Build Email context
        // -------------------
        const formDataValue = ctx.FormData as Record<string, any>; // Using FormData instead of ListSubData as Issue Type will never be touched
        const emails: EmailRouterContext = {
          status: currentStatusval === "Completed" ? "Completed" : "Cancelled",
          requesterName: flowTracker.requestor.name,
          requesterEmail: flowTracker.requestor.email,
          itemID: props.formContext.pageContext.listItem?.id,
          requestTypeText: formDataValue["Issue_x0020_Type"],
          formContext: props.formContext,
        };

        return { ok: true, value: emails };
      }

      // Invalid Status select
      case "P1900": {
        let alertMessage = "Please select correct Status to continue!";
        if (lastEntry?.stepId === "P100")
          alertMessage =
            "Invalid Status selected. Please choose Route to Submitter, Route to Team Lead, or Assigned to TW.";
        if (
          lastEntry?.stepId === "P300" ||
          lastEntry?.stepId === "P400" ||
          lastEntry?.stepId === "P800" ||
          lastEntry?.stepId === "P900" ||
          lastEntry?.stepId === "P1400" ||
          lastEntry?.stepId === "P1500"
        )
          alertMessage = "Invalid Status selected. Please choose Cancelled.";
        if (lastEntry?.stepId === "P500")
          alertMessage =
            "Invalid Status selected. Please choose Route to Submitter, Route to Team Lead, In Progress, Reassign to TW, Completed, or Cancelled.";
        if (lastEntry?.stepId === "P1000")
          alertMessage =
            "Invalid Status selected. Please choose Route to Submitter, Route to Team Lead, Reassign to TW, Completed, or Cancelled.";

        alert(alertMessage);
        return { ok: false };
      }
      //Returned from Route
      case "P600":
      case "P1200":
      case "P2000": {
        const checkComments: boolean = commentsHandler();
        if (checkComments === false) {
          alert("Please enter brief notes in Comments field!");
          return { ok: false };
        }

        addHistory({
          stepId: nextStepId,
          status: "Returned from Route",
          timestamp: new Date().toISOString(),
          modifiedby: currentUser.userFullName,
        });
        ctx.ListSubData = {
          ...ctx.ListSubData,
          Status: "Returned from Route",
        };
        // -------------------
        // Build Email context
        // -------------------
        const formDataValue = ctx.FormData as Record<string, any>; // Using FormData instead of ListSubData as Issue Type will never be touched
        const emails: EmailRouterContext = {
          status: "Returned from Route",
          routerName: lastEntry?.routerName,
          routerEmail: lastEntry?.routerEmail,
          requesterName: flowTracker.requestor.name,
          itemID: props.formContext.pageContext.listItem?.id,
          requestTypeText: formDataValue["Issue_x0020_Type"],
          formContext: props.formContext,
        };
        //If Resolution is selected accidentally, it will be removed
        delete (ctx.ListSubData as any)["Resolution"];
        return { ok: true, value: emails };
      }
      default:
        return { ok: true };
    }
  };

  const handleSubmit = async (e: React.MouseEvent<HTMLButtonElement>) => {
    e.preventDefault(); //Stop default action
    if (ctx.FormMode === FORM_MODE_DISPLAY) return; //Ignore when Form is ViewForm
    props.OnSubmit(true);
    showSpinner("Getting Info!");
    try {
      showSpinner("Resolving PeoplePicker...");
      // ✅ Wait for all async field commits (PeoplePicker registers its promise)
      const ok = await ctx.awaitAsync?.(15000);
      if (ok === false) {
        alert("PeoplePicker is still resolving. Please try again.");
        endSubmitUI();
        return;
      }

      showSpinner("Validating...."); //Update Label

      const valid = await validateForm(); //Validate Form
      if (!valid) {
        //Fail
        endSubmitUI();
        return; //stop
      }

      /* -------- Add FLOW ENGINE LOGIC HERE -------- */
      const stepHandlerOk = await stepHandler();
      if (!stepHandlerOk.ok) {
        endSubmitUI();
        return;
      }

      showSpinner("Saving...."); //Update Label
      const saveResult = await saveItem(); //Save item
      const createdNew = saveResult.status === 201; //New vs Edit
      let itemId = props.formContext.item?.ID as number; //Default ID
      if (createdNew) itemId = saveResult.data?.ID as number; //New ID
      let alertText =
        saveResult.status === 201
          ? "Thank you for Submitting Let's Fix It Request!"
          : saveResult.status === 204
          ? "Your changes are Successfully submitted!"
          : saveResult.statusText; //Base message
      if (saveResult.status === 201 || saveResult.status === 204) {
        const res: any = ctx.GlobalReturnData(); //Get Latest state
        /** Add attachments if any exists */
        if (Object.hasOwn(res?.listData ?? {}, "attachments")) {
          const settled = await uploadAttachments(
            res.listData.attachments,
            itemId
          );
          const anyFailed = settled.some((r: any) => {
            //Any failure?
            return r.status === "fulfilled" ? r.value?.status !== 200 : true; //Non-200 or rejected
          });
          if (anyFailed)
            alertText = `${alertText}, however ran into issue with attachments`; //Add warning
        }
        /** ------ SEND EMAIL ------ */
        if (stepHandlerOk.value !== undefined) {
          const allEmails: EmailRouterContext = {
            ...stepHandlerOk.value,
            itemID: itemId,
          };
          const build: EmailPayload[] | null = buildEmail(allEmails);
          if (build !== null) {
            const resultEmail: FlowResult<FlowBody> = await sendEmail(
              build,
              props.formContext
            );
            console.log(resultEmail);
          }
        }
      }

      alert(alertText); //Final message
      endSaveSuccessUI();
    } catch (error: any) {
      alert(error?.statusText ?? error);
      endSubmitUI(); // Reset UI
    }
  };

  //Re-Render when Submit button is pressed to disable the field
  React.useEffect(() => {
    setIsDisabled(props.submitting);
  }, [props.submitting]);
  React.useEffect(() => {
    //For DisplayForm button will be disabled by default
    if (ctx.FormMode === FORM_MODE_DISPLAY) {
      setIsHidden(true);
    } else {
      const decision = evaluateFieldRules(btnId, {
        formMode: ctx.FormMode,
        formData: ctx.FormData,
        curUserInfo: ctx.curUserInfo,
        formConfigJson: ctx.formRules,
      });
      if (decision.isDisabled !== undefined) {
        setIsDisabled(decision.isDisabled);
      }
      if (decision.isHidden !== undefined) {
        setIsHidden(decision.isHidden);
      } else {
        setIsHidden(false); // ✅ reset so it doesn't get stuck hidden
      }
    }
  }, []);
  return (
    <>
      <div
        className="fieldClass"
        style={{ display: isHidden ? "none" : "block", textAlign: "right" }}
      >
        <Button
          appearance="primary"
          id={btnId}
          title="Submit"
          onClick={handleSubmit}
          {...(isDisabled && { disabled: true })}
        >
          Submit
        </Button>
      </div>

      <div
        className="spinner-container"
        style={{ display: spinnerHidden ? "none" : "block" }}
      >
        <Spinner labelPosition="after" label={spinnerLabel} />
      </div>
    </>
  );

  // return (
  //   <>
  //
  //     <div
  //       className="fieldClass"
  //       style={{ display: "block", textAlign: "right" }}
  //     >
  //       <Button
  //         appearance="primary"
  //         title="Submit"
  //       >
  //         Submit
  //       </Button>
  //     </div>
  //
  //   </>
  // );
}
