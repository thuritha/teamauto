import * as React from 'react';
import styles from './TeamAutomobiles.module.scss';


import { ITeamAutomobilesProps } from './ITeamAutomobilesProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { ITeamAutomobilesState } from './ITeamAutomobilesState';
import { spOperation } from '../Services/spService';

import {
  DefaultButton,
  Dropdown,
  TextField,
  Stack,
  IStackTokens,
  IDropdownStyles,
  PrimaryButton,
  IStyleSet,
  
} from "office-ui-fabric-react";

import { Label, ILabelStyles } from "office-ui-fabric-react/lib/Label";
import {
  IPivotStyles,
  Pivot,
  PivotItem,
  PivotLinkFormat,
  PivotLinkSize,
} from "office-ui-fabric-react/lib/Pivot";

import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

const pivotStyles: Partial<IStyleSet<IPivotStyles>> = {
  link: { width: "20%" },
  linkIsSelected: { width: "20%" },
};
const stackTokens: IStackTokens = { childrenGap: 50 };
const bigVertStack: IStackTokens = { childrenGap: 20 };
const SmallVertStack: IStackTokens = { childrenGap: 20 };
const dropdownStyles: Partial<IDropdownStyles> = { dropdown: { width: 250 } };


export default class TeamAutomobiles extends React.Component<ITeamAutomobilesProps, ITeamAutomobilesState, {}> {
  
  public _spOps: spOperation;
  constructor(props: ITeamAutomobilesProps) {
    super(props);
    console.log("Constructor called!!");
    this._spOps = new spOperation();
    this.state = {
      customerNameList: [],
      productNameList: [],
      orderIdList: [],
      CustomerName: "",
      CustomerId: "",
      ProductId: "",
      ProductName: "",
      ProductUnitPrice: "",
      ProductExpiryDate: "",
      ProductType: "",
      NumberofUnits: "",
      TotalValue: "",
      orderId: "",
      status: "",
      whichButton: "Create",
    };
  }
  public componentDidMount() {
    console.log("Component Did Mount called!!");
    // getCustomerNameList
    this._spOps.getCustomerNameList(this.props.context).then((result: any) => {
      this.setState({ customerNameList: result });
      // console.log(this.state.customerNameList);
    });

    // getProductNameList
    this._spOps.getProductNameList(this.props.context).then((result: any) => {
      this.setState({ productNameList: result });
      // console.log(this.state.productNameList);
    });

    // getOrderList
    this._spOps.getOrderList(this.props.context).then((result: any) => {
      console.log(result);
      this.setState({ orderIdList: result });
    });
  }
  /**
   * getCustomerName
   * this function is called when a dropdown item is changed
   * To save the CustomerName and id to this.state
   */
  public getCustomerName = (event: any, data: any) => {
    console.log("getCustomerName called!!");
    // console.log(data);
    this.setState({
      CustomerName: data.text,
      CustomerId: data.key,
    });
  }

  /**
   * getProductName
   * this function is called when a dropdown item is changed
   * To save the productName,id,type,date and unit value to this.state
   * and show it in the form automatically
   */
  public getProductName = (event: any, data: any) => {
    console.log("getProductName called!!");
    // console.log(data);
    this._spOps
      .getProductDetails(this.props.context, data)
      .then((result: any) => {
        let date = new Date(result.ProductExpiryDate);
        var totalValue: any;
        if (this.state.NumberofUnits === "") {
          // Update Total Value when Number of Unit is not zero! otherwise don't update!!
          totalValue = this.state.TotalValue;
        } else if (this.state.NumberofUnits === "0") {
          totalValue = this.state.TotalValue;
        } else {
          totalValue =
            result.ProductUnitPrice * this.state.NumberofUnits;
        }
        // console.log(results);
        // console.log(result.ProductExpiryDate);
        // console.log(date);
        this.setState({
          ProductName: data.text,
          ProductId: data.key,
          ProductExpiryDate: date,
          ProductType: result.ProductType,
          ProductUnitPrice: result.ProductUnitPrice,
          TotalValue: totalValue,
        });
      });
  }
  /**
   * setNumberofUnits is called when number of units is changed to store the value in state.
   */
  public setNumberofUnits = (event: any, data: any) => {
    console.log("setNumberofUnits called!!");
    // console.log(this.state.ProductUnitPrice, data);
    var numberofUnits: any;
    var totalValue: any;
    if (data === "0") {
      numberofUnits = data;
      totalValue = "";
    } else if (data === "") {
      console.log(
        "setNumberofUnits called -> In ifelse -> data = '' statement!!"
      );
      numberofUnits = data;
      totalValue = "";
    } else if (this.state.ProductUnitPrice === "") {
      console.log(
        "setNumberofUnits called -> In ifelse -> UnitPrice = '' statement!!"
      );
      numberofUnits = parseInt(data);
      totalValue = "";
    } else {
      console.log("setNumberofUnits called -> In else statement!!");
      numberofUnits = parseInt(data);
      var priceofunit: number = parseInt(this.state.ProductUnitPrice);
      totalValue = numberofUnits * priceofunit;
    }
    // console.log(numberofUnits);
    // console.log(priceofunit);
    // console.log(totalValue);
    this.setState({
      NumberofUnits: numberofUnits,
      TotalValue: totalValue,
    });
    return;
  }
  /**
   * validateItemAndAdd and upload the new item
   */
  public validateItemAndAdd = () => {
    console.log("validateItemAndAdd called!!");
    let myStateList = [
      this.state.CustomerId,
      this.state.CustomerName,
      this.state.ProductId,
      this.state.ProductName,
      this.state.ProductType,
      this.state.ProductUnitPrice,
      this.state.NumberofUnits,
      this.state.TotalValue,
    ];
    // console.log(myStateList);
    for (let i = 0; i < myStateList.length; i++) {
      if (myStateList[i] === "") {
        this.setState({ status: "Fill all Details!" });
        return;
      }
    }

    console.log("Validate Complete and Uploading Order Details");
    this._spOps
      .createItems(this.props.context, this.state)
      .then((result: string) => {
        this.setState({ status: result });
      });
  }

  /**
   * validateItemAndModify
   */
  public validateItemAndModify = () => {
    if (this.state.orderId === "" || this.state.orderId === null) {
      this.setState({ status: "Enter Order Id" });
      return;
    } else {
      this._spOps.updateItem(this.state).then((status) => {
        this.setState({ status: status });
      });
    }
  }
  /**
   * validateAndDelete
   */
  public validateAndDelete = () => {
    if (this.state.orderId === "" || this.state.orderId === null) {
      this.setState({ status: "Enter Order Id" });
      return;
    } else {
      this._spOps.deleteItem(this.state.orderId).then((response) => {
        this.setState({ status: response });
      });
    }
  }
  /**
   * getOrderDetailsToUpdate is called when order id field is changed to get the item details.
   */
  public getOrderDetailsToUpdate = (event: any, data: any) => {
    // Valid Order Id -> Not empty -> not zero
    console.log("getOrderDetailsToUpdate called!");
    if (data === "") {
      console.log("Empty data");
      return;
    }
    // Now get the Order List details from rest and call setstate to change the state
    this._spOps.getUpdateitem(this.props.context, data).then((results) => {
      var result = results.value[0];
      // console.log(result);
      // find customer name
      var customerId = result.Customer_x0020_IDId;
      var customerName;
      this.state.customerNameList.forEach((item) => {
        var flag = false;
        if (item.key === customerId && flag === false) {
          customerName = item.text;
          flag = true;
        }
      });

      var productId = result.Product_x0020_IDId;
      var productName;
      this.state.productNameList.forEach((item) => {
        var flag = false;
        if (item.key === productId && flag === false) {
          productName = item.text;
          flag = true;
        }
      });

      // Now find the product details to fill
      var data1 = { key: result.Product_x0020_IDId, text: productName };
      this.getProductName({}, data1);

      this.setState({
        orderId: data.key,
        CustomerName: customerName,
        CustomerId: customerId,
        NumberofUnits: result.UnitsSold,
      });
    });
  }
  /**
   * controlTabButton is called when tabs are changed to store which tab is active.
   */
  public controlTabButton = (data: any) => {
    console.log("Tab Changed");
    console.log(data);
    if (data.props.itemKey === "1") {
      // Add tab clicked
      // reset the tab and setstate for button
      this.setState({ whichButton: "Create" });
    } else if (data.props.itemKey === "2") {
      this.setState({ whichButton: "Update" });
    } else if (data.props.itemKey === "3") {
      this.setState({ whichButton: "Delete" });
    }
  }
  /**
   * resetForm
   */
  public resetForm = () => {
    // Will reset the state of disable text field - call setstate to change state
    // Will clear text for active text field -
    console.log("resetForm called!!");
    this.setState({
      orderId: null,
      CustomerName: "",
      CustomerId: null,
      ProductId: null,
      ProductName: "",
      ProductUnitPrice: "",
      ProductExpiryDate: "",
      ProductType: "",
      NumberofUnits: "",
      TotalValue: "",
      status: "Reset Done successfully!!",
    });
    this.componentDidMount();
  }
  /**
   * renderButton is used to show active tab's button - eg: for ADD tab button should be SAVE button
   */
  public renderButton = () => {
    if (this.state.whichButton === "Create") {
      return (
        <PrimaryButton
          text="SAVE"
          onClick={this.validateItemAndAdd}
        ></PrimaryButton>
      );
    } else if (this.state.whichButton === "Update") {
      return (
        <PrimaryButton
          text="MODIFY"
          onClick={this.validateItemAndModify}
        ></PrimaryButton>
      );
    } else if (this.state.whichButton === "Delete") {
      return (
        <PrimaryButton
          text="DELETE"
          onClick={this.validateAndDelete}
          // onClick={() =>
          //   this._spOps
          //     .deleteItem(this.props.context, this.state.orderId)
          //     .then((status) => {
          //       this.setState({ status: status });
          //     })
          // }
        ></PrimaryButton>
      );
    }
  }

  
  
  public render(): React.ReactElement<ITeamAutomobilesProps> {
    return (
      <div className={ styles.teamAutomobiles }>
        <div className={ styles.container }>
          <div className={styles.row }>
            <div className={ styles.column }>
            <div className={styles.title}>Book Your Car Now</div>
              <div className={styles.subTitle}>Limited Stocks only Available</div>
              <hr></hr>
                <div>
                <Pivot styles={pivotStyles} aria-label="Basic Pivot Example" 
                                 onLinkClick={this.controlTabButton}>
              <PivotItem headerText="ADD" itemKey="1" >
                
                <div className={styles.emptyheight}></div>
              </PivotItem>
              <PivotItem headerText="UPDATE" itemKey="2">
              <div className={styles.emptyheight}></div>
                <Label>Select Order Id below.</Label>
                <Dropdown
                  required
                  selectedKey={[this.state.orderId]}
                  prefix="Order Id"
                  options={this.state.orderIdList}
                  onChange={this.getOrderDetailsToUpdate}
                ></Dropdown>
              </PivotItem>
              <PivotItem headerText="DELETE" itemKey="3" >
              <div className={styles.emptyheight}></div>
                <Label>Select Order Id below.</Label>
                <Dropdown
                  required
                  selectedKey={[this.state.orderId]}
                  prefix="Order Id"
                  options={this.state.orderIdList}
                  onChange={this.getOrderDetailsToUpdate}
                ></Dropdown>
              </PivotItem>
            </Pivot>

            <Dropdown
              required
              selectedKey={[this.state.CustomerId]}
              id="forReset1"
              label="Enter Customer Name"
              options={this.state.customerNameList}
              onChange={this.getCustomerName}
            ></Dropdown>

            <Dropdown
                  required
                  selectedKey={[this.state.ProductId]}
                  id="forReset2"
                  label="Enter Product Name"
                  options={this.state.productNameList}
                  onChange={this.getProductName}
                ></Dropdown>

            <Stack horizontal wrap tokens={stackTokens}>
            <Stack tokens={SmallVertStack}>
            <TextField
                  label="Product Type"
                  readOnly
                  placeholder={this.state.ProductType}
                />

            <TextField
                  id="forReset3"
                  label="Number of Units"
                  type="number"
                  min={0}
                  required
                  value={this.state.NumberofUnits}
                  onChange={this.setNumberofUnits}
                />
            </Stack>
            <Stack tokens={SmallVertStack}>
                <TextField
                  label="Product Expiry Date"
                  readOnly
                  placeholder={
                    this.state.ProductExpiryDate === ""
                      ? ""
                      : new Date(this.state.ProductExpiryDate).toDateString()
                  }
                />

              <div>
              <Label>Total Sales Price</Label>
              <TextField
                ariaLabel="disabled Product Sales Price"
                readOnly
                prefix="Rs. "
                placeholder={this.state.TotalValue}
              />
            </div>

            </Stack>
            <Stack tokens={SmallVertStack}>
                <TextField
                  label="Product Unit Price"
                  readOnly
                  placeholder={this.state.ProductUnitPrice}
                />
            </Stack>
            </Stack>
            
            <div className={styles.emptyheight}>{this.state.status}</div>
            <div className={styles.emptyheight}></div>
            <Stack horizontal tokens={stackTokens}>
              {this.renderButton()}
              <DefaultButton
                text="CLEAR"
                onClick={() => this.resetForm()}
              ></DefaultButton>
            </Stack>
            <div className={styles.emptyheight}></div>
                </div>

            </div>
          </div>
        </div>
      </div>
    );
  }
}
