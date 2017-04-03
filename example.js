const NodePowerBI = require('./');

const PowerBI = NodePowerBI.powerbi;
const PowerBIAuthServer = NodePowerBI.powerbiAuth;

const uuid = require('uuid');

let clientId = process.env.POWERBI_CLIENT_ID || null;
let redirectUri = process.env.POWERBI_REDIRECT_URI || 'http://localhost:3000';

let powerbiAuthServer = new PowerBIAuthServer(clientId, redirectUri);

powerbiAuthServer.init();

let userUuid = uuid();

console.log(powerbiAuthServer.generateAuthURL(userUuid));

powerbiAuthServer.setCallbackOnAuthentication(userUuid, (userInfo)=>{
  let accessToken = userInfo['accessToken'];
  let refreshToken = userInfo['refreshToken'];
  let powerbi = new PowerBI(accessToken, refreshToken);

  powerbi.deleteRows('21892f21-ea66-495a-b2a5-5c868b3a1e05', 'ExampleTable', (err, result)=>{
    console.log(err);
    console.log(result);
  });

  /*powerbi.addRows('21892f21-ea66-495a-b2a5-5c868b3a1e05', 'ExampleTable', [{ExampleColumn: Math.random()}], (err, result)=>{
    console.log(result);
  });*/

  /*powerbi.listTables('21892f21-ea66-495a-b2a5-5c868b3a1e05', (err, tables)=>{
    console.log(tables);
  });*/

  /*powerbi.createDataset('ExampleSLR', 'Push', [{ name: 'ExampleTable', columns: [{name: 'ExampleColumn', dataType: 'String'}]}], (err, dataset)=>{
    if(err){
      console.log(err);
    }
    else{
      console.log('Created dataset');
    }
    /*powerbi.deleteDataset(dataset.id, (err)=>{
      if(err){
        console.log(err);
      }
      else{
        console.log('Deleted dataset');
      }
    });
  });*/
});
