'use strict';

const request = require('request');
const express = require('express');

class PowerBI {
  constructor(accessToken) {
    this.baseURI = 'https://api.powerbi.com/v1.0/myorg';
    this.accessToken = accessToken;
  }

  listDatasets(callback) {
    request({
      url: this.baseURI + '/datasets',
      method: 'GET',
      headers: {
        Authorization: 'Bearer ' + this.accessToken,
        'Content-Length': 0
      }
    }, (error, response, body) => {
      callback(JSON.parse(body).value);
    });
  }

  getDatasetByName(name, callback) {
    this.listDatasets(datasets => {
      for (let i = 0; i < datasets.length; i++) {
        let dataset = datasets[i];
        if (dataset.name === name) {
          callback(null, dataset);
          return;
        }
      }
      callback('Dataset ' + name + ' not found');
    });
  }

  createDataset(name, defaultMode, tables, callback) {
    let datasetRequestBody = {
      name: name,
      defaultMode: defaultMode,
      tables: tables
    };

    request({
      method: 'POST',
      url: this.baseURI + '/datasets?defaultRetentionPolicy=None',
      headers: {
        'Content-Type': 'application/json',
        Authorization: 'Bearer ' + this.accessToken
      },
      body: JSON.stringify(datasetRequestBody)
    }, (err, res, body) => {
      if (err) {
        callback(err);
      } else {
        callback(null, JSON.parse(body));
      }
    });
  }

  deleteDataset(id, callback) {
    request({
      method: 'DELETE',
      url: this.baseURI + '/datasets/' + id,
      headers: {
        'Content-Type': 'application/json',
        Authorization: 'Bearer ' + this.accessToken
      }
    }, (err, response) => {
      if (err) {
        callback(err);
      } else if (response.statusCode === 200) {
        callback(null);
      } else {
        callback('Unknown error while deleting dataset: ' + id);
      }
    });
  }

  listTables(datasetId, callback) {
    request({
      method: 'GET',
      url: this.baseURI + '/datasets/' + datasetId + '/tables',
      headers: {
        'Content-Type': 'application/json',
        Authorization: 'Bearer ' + this.accessToken
      }
    }, (err, response, body) => {
      // Check if error
      if (err) {
        callback(err);
      } else if (response.statusCode === 200) {
        let response = JSON.parse(body);
        if (response.error) {
          callback(response);
        } else {
          callback(null, response.value);
        }
      } else {
        callback('Unknown error while retrieving tables of dataset: ' + datasetId);
      }
    });
  }

  addRows(datasetId, tableName, rows, callback) {
    let rowsWrapper = {rows: rows};
    request({
      method: 'POST',
      url: this.baseURI + '/datasets/' + datasetId + '/tables/' + tableName + '/rows',
      headers: {
        'Content-Type': 'application/json',
        Authorization: 'Bearer ' + this.accessToken
      },
      body: JSON.stringify(rowsWrapper)
    }, (err, response, body) => {
      if (err) {
        callback(err);
      } else {
        callback(null, JSON.parse(body));
      }
    });
  }

  deleteRows(datasetId, tableName, callback) {
    request({
      method: 'DELETE',
      url: this.baseURI + '/datasets/' + datasetId + '/tables/' + tableName + '/rows'
    }, (err, response) => {
      if (err) {
        callback(err);
      } else if (response.statusCode === 200) {
        callback(err);
      } else {
        callback(null, true);
      }
    });
  }
}

class PowerBIAuthServer {
  constructor(clientId, redirectUri, clientSecret) {
    this.users = {}; // Temporal variable to store people who is authenticating
    this.clientId = clientId;
    this.redirectUri = redirectUri;
    this.clientSecret = clientSecret;
    this.callback = () => {};
  }

  init() {
    // Create a server to listen redirect in the first step of authentication
    this.app = express();
    this.app.get('/', (req, res) => {
      let authorizationCode = req.query.code;
      let userUuid = req.query.state;

      // Prepare request body
      let requestBody = {
        grant_type: 'authorization_code',
        client_id: this.clientId,
        code: authorizationCode,
        redirect_uri: this.redirectUri,
        resource: 'https://analysis.windows.net/powerbi/api'
      };

      if (this.clientSecret) {
        requestBody.client_secret = this.clientSecret;
      }

      // Send the request to the oauth server and retrieve tokens
      request({
        url: 'https://login.microsoftonline.com/common/oauth2/token/',
        method: 'POST',
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded'
        },
        form: requestBody
      }, (err, response, body) => {
        if (err) {
          this.users[userUuid].callback(err);
        } else {
          let parsedBody = JSON.parse(body);
          this.users[userUuid].accessToken = parsedBody.access_token;
          this.users[userUuid].refreshToken = parsedBody.refresh_token;
          // Return data
          this.users[userUuid].callback(null, this.users[userUuid]);
          this.users[userUuid] = {}; // Remove
        }
      });
      res.send('Correctly authenticated, you can close this window');
    });

    this.app.listen(3000, () => {
      console.log('PowerBI auth server listening on port 3000');
    });
  }

  generateAuthURL(userUuid) {
    let url = 'https://login.windows.net/common/oauth2/authorize?response_type=code&resource=https://analysis.windows.net/powerbi/api&';
    // Set client ID
    if (this.clientId) {
      url += 'client_id=' + this.clientId + '&';
    }
    // Set redirect uri
    if (this.redirectUri) {
      url += 'redirect_uri=' + this.redirectUri + '&';
    } else {
      throw new Error('Redirect URI required');
    }
    url += 'state=' + userUuid;
    return url;
  }

  setCallbackOnAuthentication(userUuid, callback) {
    this.users[userUuid] = {};
    this.users[userUuid].callback = callback;
  }

}

module.exports = {powerbi: PowerBI, powerbiAuth: PowerBIAuthServer};
