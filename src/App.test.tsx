import React from 'react';
import { render, screen } from '@testing-library/react';
import { PublicClientApplication } from '@azure/msal-browser';
import App from './App';
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
    msalInstance.setActiveAccount(accounts[0]);
}

test('renders learn react link', () => {
    render(<App  pca = { msalInstance }/>);
    const linkElement = screen.getByText(/learn react/i);
    expect(linkElement).toBeInTheDocument();
});
