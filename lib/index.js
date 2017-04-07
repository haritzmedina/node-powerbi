'use strict';

const request = require('request');
const restify = require('restify');

class PowerBI {
  constructor(accessToken, refreshToken) {
    this.baseURI = 'https://api.powerbi.com/v1.0/myorg';
    this.accessToken = accessToken;
    this.refreshToken = refreshToken;
  }

  isValidAccessToken(callback){
    this.listDatasets((err, result)=>{
      if(err){
        callback(false);
      }
      else{
        // TODO Check if result is access token valid
        callback(true);
      }
    })
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
      if(error){
        callback(error);
      }
      else{
        callback(null, JSON.parse(body).value);
      }
    });
  }

  getDatasetByName(name, callback) {
    this.listDatasets((err, datasets) => {
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
        callback(null);
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
  constructor(options, restifyServer){
    this.users = {}; // Temporal variable to store people who is authenticating
    this.clientId = options.clientId;
    this.redirectUri = options.redirectUri || 'http://localhost:3000';
    this.clientSecret = options.clientSecret || null;
    this.serverPort = options.serverPort || 3000;
    this.restifyEndpoint = options.restifyEndpoint || '/';
    if(restifyServer){
      this.restifyServer = restifyServer;
    }
    this.callback = () => {};
  }

  init() {
    // If server is not created, create it
    if(!this.restifyServer){
      this.restifyServer = restify.createServer({});
      this.restifyServer.listen(this.serverPort, ()=>{
        console.log('PowerBI auth server listening on port '+this.serverPort);
      });
    }

    // To retrieve query params '/?code=sadksad&state=sada, restify requires this flag
    this.restifyServer.use(restify.queryParser());

    // Define get response a server to listen redirect in the first step of authentication
    this.restifyServer.get(this.restifyEndpoint, (req, res, next)=>{
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
        res.send('Correctly authenticated, you can close this window');
      });
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

  refreshAccessToken(refreshToken, callback){
    let requestBody = {
      client_id: this.clientId,
      refresh_token: refreshToken,
      grant_type: 'refresh_token',
      resource: 'https://analysis.windows.net/powerbi/api',
      client_secret: this.clientSecret
    };

    request({
      url: 'https://login.microsoftonline.com/common/oauth2/token/',
      method: 'POST',
      headers:  {
        'Content-Type': 'application/x-www-form-urlencoded'
      },
      form: requestBody
    }, (err, response, body) => {
      if(err){
        callback(err);
      }
      else{
        let parsedBody = JSON.parse(body);
        callback(null, {accessToken: parsedBody.access_token, refreshToken: parsedBody.refresh_token});
      }
    });
  }

}

module.exports = {powerbi: PowerBI, powerbiAuth: PowerBIAuthServer};
