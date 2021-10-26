import React from 'react';
import ReactDOM from 'react-dom';
import './index.css';
import App from './App';
import reportWebVitals from './reportWebVitals';
import {
  PublicClientApplication,
  EventType,
  EventMessage,
  AuthenticationResult } from '@azure/msal-browser';
import config from './Config';

const msalInstance = new PublicClientApplication({
  auth: {
    clientId: config.appId,
    redirectUri: config.redirectUri
  },
  cache: {
    cacheLocation: 'sessionStorage',
    storeAuthStateInCookie: true
  }
});

const accounts = msalInstance.getAllAccounts();
if (accounts && accounts.length > 0) {

  debugger;

  msalInstance.setActiveAccount(accounts[0]);
}

msalInstance.addEventCallback((event: EventMessage) => {
  if (event.eventType === EventType.LOGIN_SUCCESS && event.payload) {

    debugger;

    const authResult = event.payload as AuthenticationResult;
    msalInstance.setActiveAccount(authResult.account);
  }
});

ReactDOM.render(
  <React.StrictMode>
    <App pca = { msalInstance }/>
  </React.StrictMode>,
  document.getElementById('root')
);

reportWebVitals();
