import "azure-devops-ui/Core/override.css";
import "./index.scss";

import * as SDK from "azure-devops-extension-sdk";
import { FormItem } from "azure-devops-ui/FormItem";
import { TextField } from "azure-devops-ui/TextField";
import * as React from "react";
import * as ReactDOM from "react-dom";

export class WorkItemControlComponent extends React.Component {
  constructor(props: {}) {
    super(props);
  }

  public async componentDidMount() {
    SDK.init();
  }

  public render(): JSX.Element {
    return (
      <>
        <FormItem label="Foo 1" className="foo-form-item">
          <TextField className="foo-text-field" value="Bar 1" readOnly={true} />
        </FormItem>
        <FormItem label="Foo 2" className="foo-form-item">
          <TextField className="foo-text-field" value="Bar 2" readOnly={true} />
        </FormItem>
      </>
    );
  }
}

ReactDOM.render(<WorkItemControlComponent />, document.getElementById("root"));
