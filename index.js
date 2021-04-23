const express = require('express');
const app = express();
const port = 5555;
const FormData = require('form-data');
const axios = require('axios');
const fs = require('fs');

const client_id = '93dd1182-113e-473f-86c5-c2d18ed3ce5b@1f546c63-7abe-44cc-8ace-05925fc8c22b';
const client_secret = 'kxxtSeCmZg9zbvpLgaUj5hAqY2RxLOJ6xC9lHIm6tRQ=';
const tenant = 'omnicon';
const tenantID = '1f546c63-7abe-44cc-8ace-05925fc8c22b';
const resource = '00000003-0000-0ff1-ce00-000000000000/' + tenant + '.sharepoint.com@' + tenantID;



app.get('/', function (req, res) {
  res.send(new Response(200, 'adios mundo', null));
});
app.listen(port, () => {
  console.log(`on http://localhost:${port}`);
  getVersionsFile()

})

async function getAccessToken() {
  try {
    let url = 'https://accounts.accesscontrol.windows.net/1f546c63-7abe-44cc-8ace-05925fc8c22b/tokens/OAuth/2/';
    let form = new FormData();
    form.append('grant_type', 'client_credentials');
    form.append('client_id', '93dd1182-113e-473f-86c5-c2d18ed3ce5b@1f546c63-7abe-44cc-8ace-05925fc8c22b');
    form.append('client_secret', 'kxxtSeCmZg9zbvpLgaUj5hAqY2RxLOJ6xC9lHIm6tRQ=');
    form.append('resource', '00000003-0000-0ff1-ce00-000000000000/omnicon.sharepoint.com@1f546c63-7abe-44cc-8ace-05925fc8c22b');
    let response = await axios.post(
      url,
      form,
      {
        headers: {
          'Accept': 'application/json;odata=verbose',
          'Content-Type': 'multipart/form-data; boundary=' + form.getBoundary(),
          'Connection': 'keep-alive',
          'Content-Length': form.getLengthSync()
        }
      }
    );
    return response.data.access_token;
  } catch (error) {
    console.log(error);
    return null;
  }
}


async function getFile() {
  let tk = await getAccessToken();
  let url = "https://omnicon.sharepoint.com/sites/KairosTest-Sharepoint/_api/web/GetFileByServerRelativeUrl('/sites/KairosTest-Sharepoint/Documentos compartidos/Kairos/proyecto de prueba 1/excel de prueba.xlsx')/";
  let response = await axios.get(
    url,
    {
      headers: {
        'Authorization': 'Bearer ' + tk,
        'Accept': 'application/json;odata=verbose',
      }
    }
  );
  console.log(response.data.d);
}


async function uploadFile() {
  let tk = await getAccessToken();
  let url = "https://omnicon.sharepoint.com/sites/KairosTest-Sharepoint/_api/web/GetFolderByServerRelativeUrl('/sites/KairosTest-Sharepoint/Documentos compartidos/Kairos/proyecto de prueba 1')/Files/add(url='pruebaJhoan.docx',overwrite=true)";
  let form = new FormData();
  fs.readFile('./documento1.docx', async function (err, data) {
    try {
      if (err) {
        throw err;
      }
      form.append('file', data, 'documento1.docx');
      let response = await axios.post(
        url,
        form,
        {
          headers: {
            'Authorization': 'Bearer ' + tk,
            'Content-Type': 'multipart/form-data; boundary=' + form.getBoundary(),
            'Content-Length': form.getLengthSync()
          }
        }
      );
      console.log(response);
    } catch (error) {
      console.log(error);
    }
  })
}


async function getVersionsFile() {
  let tk = await getAccessToken();
  let url = "https://omnicon.sharepoint.com/sites/KairosTest-Sharepoint/_api/web/GetFileByServerRelativeUrl('/sites/KairosTest-Sharepoint/Documentos compartidos/Kairos/proyecto de prueba 1/pruebaJhoan.docx')/versions";
  let response = await axios.get(
    url,
    {
      headers: {
        'Authorization': 'Bearer ' + tk,
        'Accept': 'application/json;odata=verbose',
      }
    }
  );
  console.log(response.data.d.results);
}


async function downloadVersionFile() {try{
  let tk = await getAccessToken();
  let url = "https://omnicon.sharepoint.com/sites/KairosTest-Sharepoint/_api/web/GetFileByServerRelativeUrl('/sites/KairosTest-Sharepoint/Documentos compartidos/Kairos/proyecto de prueba 1/pruebaJhoan.docx')/versions(512)/$value"
  let response = await axios({
    url: url,
    method: 'GET',
    headers: {
      'Authorization': 'Bearer ' + tk,
      'Accept': 'application/json;odata=verbose',
    }
    , responseType: 'stream'
  }
  );
  let writer = fs.createWriteStream('version1.docx');
  response.data.pipe(writer)
  return new Promise((resolve, reject) => {
    writer.on('finish', resolve);
    writer.on('error', reject)
  })}catch(error){
    console.log(error);
  }
}


// /sites/KairosTest-Sharepoint/Documentos compartidos/Kairos/proyecto de prueba 1/pruebaJhoan.docx'