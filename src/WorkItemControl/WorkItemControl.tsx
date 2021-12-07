import "azure-devops-ui/Core/override.css";
import "./index.scss";

import {
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

enum FieldReferenceName {
  CreatedDate = "System.CreatedDate",
  ActivatedDate = "Microsoft.VSTS.Common.ActivatedDate",
  ClosedDate = "Microsoft.VSTS.Common.ClosedDate",
}

interface TimeMetricsProps {}

interface TimeMetricsState {
  dateCreated: Date | undefined;
  dateActivated: Date | undefined;
  dateClosed: Date | undefined;
}

export class TimeMetricsComponent extends React.Component<TimeMetricsProps, TimeMetricsState> {
  constructor(props: TimeMetricsProps) {
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

  public componentWillUnmount() {
    SDK.unregister(SDK.getContributionId());
  }

  public render(): JSX.Element {
    var { dateCreated, dateActivated, dateClosed } = this.state;

    return (
      <>
        <FormItem label="Lead Time" className="time-metrics-form-item">
          <TextField
            className="time-metrics-text-field"
            value={this.getDuration(dateCreated, dateClosed)}
            readOnly={true}
          />
        </FormItem>
        <FormItem label="Cycle Time" className="time-metrics-form-item">
          <TextField
            className="time-metrics-text-field"
            value={this.getDuration(dateActivated, dateClosed)}
            readOnly={true}
          />
        </FormItem>
      </>
    );
  }

  private getDateOrUndefined(valueFromPrevState: Date | undefined, valueFromSdk: any): Date | undefined {
    if (valueFromSdk === null || valueFromSdk === undefined) {
      return undefined;
    }
    /*
      Scenario: Begin @ New, Update → Committed, Save
        dateActivated = Object { type: 1, toString: key(), valueOf: key() }

      Scenario: Begin @ Comitted, Update → Done, Save
        dateClosed = Object { type: 1, toString: key(), valueOf: key() }
    */
    if (typeof valueFromSdk === "object" && "type" in valueFromSdk) {
      // TODO Better, TypeScript-y way to perform this check?
      return valueFromPrevState ?? new Date();
    }
    return valueFromSdk as Date;
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

  private async getStateFromSdk(prevState: Readonly<TimeMetricsState>): Promise<TimeMetricsState> {
    const service = await SDK.getService<IWorkItemFormService>(
      WorkItemTrackingServiceIds.WorkItemFormService
    );
    const datesFromSdk = await service.getFieldValues(
      [FieldReferenceName.CreatedDate, FieldReferenceName.ActivatedDate, FieldReferenceName.ClosedDate],
      {
        returnOriginalValue: false,
      }
    );
    console.dir(datesFromSdk);
    return {
      dateCreated: this.getDateOrUndefined(
        prevState.dateCreated,
        datesFromSdk[FieldReferenceName.CreatedDate]
      ),
      dateActivated: this.getDateOrUndefined(
        prevState.dateActivated,
        datesFromSdk[FieldReferenceName.ActivatedDate]
      ),
      dateClosed: this.getDateOrUndefined(prevState.dateClosed, datesFromSdk[FieldReferenceName.ClosedDate]),
    };
  }

  private registerEvents(): void {
    SDK.register(SDK.getContributionId(), () => {
      return {
        onLoaded: (args: IWorkItemLoadedArgs) => {
          if (args.isNew) {
            return;
          }
          this.getStateFromSdk(this.state).then((state) => {
            this.setState(state);
          });
        },

        onSaved: () => {
          this.getStateFromSdk(this.state).then((state) => {
            this.setState(state);
          });
        },
      };
    });
  }
}

ReactDOM.render(<TimeMetricsComponent />, document.getElementById("root"));
