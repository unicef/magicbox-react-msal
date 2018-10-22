import * as Msal from 'msal';
export default class AuthService {
  constructor() {
    let PROD_REDIRECT_URI = process.env.REACT_APP_REPLY_URL;
    let redirectUri = window.location.origin;
    if (window.location.hostname !== '127.0.0.1') {
      redirectUri = PROD_REDIRECT_URI;
    }
    this.applicationConfig = {
      clientID: process.env.REACT_APP_CLIENT_ID,
      authority: process.env.REACT_APP_AUTHORITY,
      graphScopes: ['user.read']
    };
    this.app = new Msal.UserAgentApplication(
      this.applicationConfig.clientID,
      '',
      () => {
        // callback for login redirect
      },
      {
        redirectUri
      }
    );
  }
  login = () => {
    return this.app.loginPopup(this.applicationConfig.graphScopes).then(
      idToken => {
        console.log(idToken)
        const user = this.app.getUser();
        if (user) {
          return user;
        } else {
          return null;
        }
      },
      () => {
        return null;
      }
    ).catch(err => {
      console.log(err)
    })
  };
  logout = () => {
    this.app.logout();
  };
  getToken = () => {
    return this.app.acquireTokenSilent(this.applicationConfig.graphScopes).then(
      accessToken => {
        return accessToken;
      },
      error => {
        return this.app
          .acquireTokenPopup(this.applicationConfig.graphScopes)
          .then(
            accessToken => {
              return accessToken;
            },
            err => {
              console.error(err);
            }
          );
      }
    );
  };
}
