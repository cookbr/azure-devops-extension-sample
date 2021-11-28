import "azure-devops-ui/Core/override.css";
import "./index.scss";

import {
  IWorkItemChangedArgs,
  IWorkItemFieldChangedArgs,
  IWorkItemFormService,
  IWorkItemLoadedArgs,
  WorkItemTrackingServiceIds,
} from "azure-devops-extension-api/WorkItemTracking";
import * as SDK from "azure-devops-extension-sdk";
import { FormItem } from "azure-devops-ui/FormItem";
import { TextField } from "azure-devops-ui/TextField";
import { Duration } from "azure-devops-ui/Duration";
import * as React from "react";
import * as ReactDOM from "react-dom";

enum WorkItemFieldReferenceNames {
  CreatedDate = "System.CreatedDate",
  ActivatedDate = "Microsoft.VSTS.Common.ActivatedDate",
  ClosedDate = "Microsoft.VSTS.Common.ClosedDate",
}

interface WorkItemFormGroupComponentState {
  dateCreated: Date | undefined;
  dateActivated: Date | undefined;
  dateCompleted: Date | undefined;
}

export class WorkItemControlComponent extends React.Component<
  {},
  WorkItemFormGroupComponentState
> {
  constructor(props: {}) {
    super(props);
    this.state = {
      dateCreated: undefined,
      dateActivated: undefined,
      dateCompleted: undefined,
    };
  }

  public componentDidMount() {
    SDK.init().then(() => {
      this.registerEvents();
    });
  }

  public render(): JSX.Element {
    // TODO Better UI when work item is active or yet to be activated
    var { dateCreated, dateActivated, dateCompleted } = this.state;

    return (
      <>
        <FormItem label="Lead Time" className="foo-form-item">
          <Duration
            startDate={dateCreated ?? new Date()}
            endDate={dateCompleted}
          />
        </FormItem>
      </>
    );
  }

  private refresh(): void {
    SDK.getService<IWorkItemFormService>(
      WorkItemTrackingServiceIds.WorkItemFormService
    )
      .then((service) => {
        const fieldReferenceNames: string[] = [
          WorkItemFieldReferenceNames.CreatedDate,
          WorkItemFieldReferenceNames.ActivatedDate,
          WorkItemFieldReferenceNames.ClosedDate,
        ];
        return service.getFieldValues(fieldReferenceNames, {
          returnOriginalValue: false,
        });
      })
      .then((dates) => {
        console.dir(dates);
        this.setState({
          dateCreated: dates[WorkItemFieldReferenceNames.CreatedDate] as
            | Date
            | undefined,
          dateActivated: dates[WorkItemFieldReferenceNames.ActivatedDate] as
            | Date
            | undefined,
          dateCompleted: dates[WorkItemFieldReferenceNames.ClosedDate] as
            | Date
            | undefined,
        });
      });
  }

  private registerEvents(): void {
    SDK.register(SDK.getContributionId(), () => {
      return {
        // Called when the active work item is modified
        onFieldChanged: (args: IWorkItemFieldChangedArgs) => {
          console.debug(`onFieldChanged - ${JSON.stringify(args)}`);
          this.refresh();
        },

        // Called when a new work item is being loaded in the UI
        onLoaded: (args: IWorkItemLoadedArgs) => {
          console.debug(`onLoaded - ${JSON.stringify(args)}`);
          this.refresh();
        },

        // Called when the active work item is being unloaded in the UI
        onUnloaded: (args: IWorkItemChangedArgs) => {
          console.debug(`onUnloaded - ${JSON.stringify(args)}`);
        },

        // Called after the work item has been saved
        onSaved: (args: IWorkItemChangedArgs) => {
          console.log(`onSaved - ${JSON.stringify(args)}`);
        },

        // Called when the work item is reset to its unmodified state (undo)
        onReset: (args: IWorkItemChangedArgs) => {
          console.log(`onReset - ${JSON.stringify(args)}`);
          this.refresh();
        },

        // Called when the work item has been refreshed from the server
        onRefreshed: (args: IWorkItemChangedArgs) => {
          console.debug(`onRefreshed - ${JSON.stringify(args)}`);
          this.refresh();
        },
      };
    });
  }
}

ReactDOM.render(<WorkItemControlComponent />, document.getElementById("root"));
