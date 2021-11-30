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
import { duration } from "azure-devops-ui/Utilities/Date";
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
  dateClosed: Date | undefined;
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
      dateClosed: undefined,
    };
  }

  public componentDidMount() {
    SDK.init().then(() => {
      this.registerEvents();
    });
  }

  public render(): JSX.Element {
    var { dateCreated, dateActivated, dateClosed } = this.state;

    return (
      <>
        <FormItem label="Lead Time" className="foo-form-item">
          <TextField
            className="foo-text-field"
            value={this.getDuration(dateCreated, dateClosed)}
            readOnly={true}
          />
        </FormItem>
        <FormItem label="Cycle Time" className="foo-form-item">
          <TextField
            className="foo-text-field"
            value={this.getDuration(dateActivated, dateClosed)}
            readOnly={true}
          />
        </FormItem>
      </>
    );
  }

  private getDuration(beginning: Date | undefined, end: Date | undefined): string {
    if (beginning && end) {
      return duration(beginning, end);
    }
    if (beginning) {
      return `TBD (Currently ${duration(beginning, end)})`;
    }
    return "";
  }

  private refresh(): void {
    SDK.getService<IWorkItemFormService>(WorkItemTrackingServiceIds.WorkItemFormService)
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
          dateCreated: dates[WorkItemFieldReferenceNames.CreatedDate] as Date | undefined,
          dateActivated: dates[WorkItemFieldReferenceNames.ActivatedDate] as
            | Date
            | undefined,
          dateClosed: dates[WorkItemFieldReferenceNames.ClosedDate] as Date | undefined,
        });
      });
  }

  private registerEvents(): void {
    // TODO Update state using *Args to forgo running through IWorkItemFormService (assuming there's a network tax)
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
