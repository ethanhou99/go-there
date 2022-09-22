// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React from 'react';
import './App.css';
import './Tab.css';
import img from '../maprt.jpg';
import { TeamsFx } from "@microsoft/teamsfx";
import { Button, Table } from "@fluentui/react-northstar"

import { Providers, ProviderState } from '@microsoft/mgt-element';
import { PeoplePicker, Person, PersonViewType, PersonCardInteraction } from '@microsoft/mgt-react';
import { TeamsFxProvider } from '@microsoft/mgt-teamsfx-provider';
import { CacheService } from '@microsoft/mgt';
import ZoomImage from './ZoomImage';

class Tab extends React.Component {

  constructor(props) {
    super(props);
    CacheService.clearCaches();

    this.state = {
      showLoginPage: undefined,
      selectedPeople: undefined,
      tableHeader:['Name', 'Email', 'User Principal Name', 'Location'],
      tableRows:[],
      csvData: []
    }
  }

  async componentDidMount() {
    await this.initTeamsFx();
    await this.initGraphToolkit(this.teamsfx, this.scope);
    await this.checkIsConsentNeeded();
  }

  async initGraphToolkit(teamsfx, scope) {
    const provider = new TeamsFxProvider(teamsfx, scope)
    Providers.globalProvider = provider;
  }

  async initTeamsFx() {
    this.teamsfx = new TeamsFx();

    // Only these two permission can be used without admin approval in microsoft tenant
    this.scope = [
      "User.Read",
      "User.ReadBasic.All",
    ];
  }

  async loginBtnClick() {
    try {
      await this.teamsfx.login(this.scope);
      Providers.globalProvider.setState(ProviderState.SignedIn);
      this.setState({
        showLoginPage: false
      });
    } catch (err) {
      if (err.message?.includes("CancelledByUser")) {
        const helpLink = "https://aka.ms/teamsfx-auth-code-flow";
        err.message += 
          "\nIf you see \"AADSTS50011: The reply URL specified in the request does not match the reply URLs configured for the application\" " + 
          "in the popup window, you may be using unmatched version for TeamsFx SDK (version >= 0.5.0) and Teams Toolkit (version < 3.3.0) or " +
          `cli (version < 0.11.0). Please refer to the help link for how to fix the issue: ${helpLink}` ;
      }

      alert("Login failed: " + err);
      return;
    }
  }

  async checkIsConsentNeeded() {
    let consentNeeded = false;
    try {
      await this.teamsfx.getCredential().getToken(this.scope);
    } catch (error) {
      consentNeeded = true;
    }
    this.setState({
      showLoginPage: consentNeeded
    });
    Providers.globalProvider.setState(consentNeeded ? ProviderState.SignedOut : ProviderState.SignedIn);
    return consentNeeded;
  }
  
  render() {

    const handleInputChange = (e) => {
      this.setState({
        selectedPeople: e.target.selectedPeople
      })

      const rows = e.target.selectedPeople.map((person, index) => {
        return {
          key: index,
          truncateContent: true,
          items: [
            {
              content: <Person userId={person.id} view={PersonViewType.oneline} personCardInteraction={PersonCardInteraction.hover}></Person>,
              truncateContent: true,
              title: person.displayName,
            },
            {
              content: person.mail,
              truncateContent: true,
              title: person.mail
            },
            {
              content: person.userPrincipalName,
              title: person.userPrincipalName,
              truncateContent: true,
            },
            {
              content: 'SHA-ZIZHU-BLD1/1707', //hardcoded
              title: person.officeLocation,
              truncateContent: true,
            },
          ],
        };
      });

      this.setState({
        tableRows: rows
      });
    };
    
    return (
      <div>
        {this.state.showLoginPage === false && <div className="flex-container">
          <div className="features-col">
            <div className="features">

              <div className="header">
                <div className="title">
                  <h2>Campus Map</h2>
                </div>
              </div>

              <div className="people-picker-area">
                <PeoplePicker userType="user" transitiveSearch="true" selectionChanged={handleInputChange} placeholder="Typing name to search people to view their campus location"></PeoplePicker>
              </div>
              <div className="table-area">
                <Table  variables={{cellContentOverflow: 'none'}} header={this.state.tableHeader} rows={this.state.tableRows} aria-label="Static table" />
              </div>
              <div className="map">
                {this.state.selectedPeople != undefined ? <ZoomImage image={img}></ZoomImage> : <h2 className="Welcome">Thank you for using Go There campus map</h2>}
              </div>
            </div>
          </div>
        </div>}

        {this.state.showLoginPage === true && <div className="auth">
          <h2>Welcome to Contact Exporter App!</h2>
          <Button primary onClick={() => this.loginBtnClick()}>Start</Button>
        </div>}
      </div>
    );
  }
}
export default Tab;
