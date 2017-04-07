const NodePowerBI = require('./');

const restify = require('restify');

const PowerBI = NodePowerBI.powerbi;
const PowerBIAuthServer = NodePowerBI.powerbiAuth;

const uuid = require('uuid');


let serverParams = {
  clientId: process.env.POWERBI_CLIENT_ID || null,
  redirectUri : process.env.POWERBI_REDIRECT_URI || 'http://localhost:3000',
  serverPort : process.env.POWERBI_SERVER_PORT || 3000,
  clientSecret: null
};

let EXAMPLE_DATASET_ID = process.env.EXAMPLE_DATASET_ID; // FIX IT to run the example

const server = restify.createServer({});
server.listen(3000, function(){
  console.log(`${server.name} listening to ${server.url}`);
});


let powerbiAuthServer = new PowerBIAuthServer(serverParams, server);

powerbiAuthServer.init();

let userUuid = uuid();

console.log(powerbiAuthServer.generateAuthURL(userUuid));

powerbiAuthServer.setCallbackOnAuthentication(userUuid, (err, userInfo) => {
  if (err) {
    console.log(err);
  }
  let accessToken = userInfo.accessToken;
  let refreshToken = userInfo.refreshToken;
  let powerbi = new PowerBI(accessToken, refreshToken);

  // EXAMPLE 0 Refresh token
  powerbiAuthServer.refreshAccessToken(refreshToken, (err, result)=>{
    console.log('Your old access token is: '+accessToken);
    console.log('Your new access token is: '+result.accessToken);
    console.log('Your old refresh token is: '+refreshToken);
    console.log('Your new refresh token is: '+result.refreshToken);
  });

  // EXAMPLE 1 Create a dataset
  powerbi.createDataset('Example', 'Push', [{name: 'ExampleTable', columns: [{name: 'ExampleColumn', dataType: 'String'}]}], (err, dataset) => {
    if (err) {
      console.log(err);
    } else {
      console.log('Dataset created with id: ' + dataset.id);
    }
  });

  powerbi.getDatasetByName('Example', (err, dataset) => {
    if (err) {
      console.log(err);
    } else {
      console.log(dataset);
    }
  });

  // EXAMPLE 2 DELETE A DATASET
  powerbi.deleteDataset(EXAMPLE_DATASET_ID, err => {
    if (err) {
      console.log(err);
    } else {
      console.log('Deleted dataset');
    }
  });

  // EXAMPLE 3 List Datasets
  powerbi.listTables(EXAMPLE_DATASET_ID, (err, tables) => {
    if (err) {
      console.log(err);
    } else {
      console.log(tables);
    }
  });

  // EXAMPLE 4 Add row to a table
  powerbi.addRows(EXAMPLE_DATASET_ID, 'ExampleTable', [{ExampleColumn: Math.random()}], (err, result) => {
    if (err) {
      console.log(err);
    } else {
      if(result){
        console.log('Correctly added');
      }
    }
  });

  // EXAMPLE 5 Remove rows from a table
  powerbi.deleteRows(EXAMPLE_DATASET_ID, 'ExampleTable', (err, result) => {
    console.log(err);
    console.log(result);
  });

});
