import querystring from "querystring";
import { Stream } from "stream";
import { Promise } from "es6-promise";

import Express from "express";
import MicrosoftGraph = require("@microsoft/microsoft-graph-client");

const accessToken = process.env.ACCESS_TOKEN || "";
const clientID = process.env.CLIENT_ID || "";
const port = 3000;


const app = Express();

app.get('/', (req, res) => res.send('OneNote Export'));

app.get('/login', (req, res) => {
  const qs = querystring.stringify({
    client_id: clientID,
    scope: "Notes.Read Notes.Read.All",
    response_type: "token",
    redirect_uri: `http://localhost:${port}/token`,
  });
  res.redirect(`https://login.microsoftonline.com/common/oauth2/v2.0/authorize?${qs}`);
});

app.get('/token', (req, res) => {
  res.status(200).send();
});

app.get('/notebooks', (req, res) => {
  const graph = createGraphClient();
  getAll(graph, graph
    .api('/me/onenote/notebooks')
    .version('v1.0')
    .orderby('displayName')
  ).then(data => sendJSON(res, data));
});

app.get('/sections', (req, res) => {
  const graph = createGraphClient();
  getAll(graph, graph
    .api('/me/onenote/sections')
    .version('v1.0')
    .orderby('displayName')
  ).then(data => sendJSON(res, data));
});

app.get('/pages', (req, res) => {
  const graph = createGraphClient();
  getAll(graph, graph
    .api('/me/onenote/pages')
    .version('v1.0')
    .orderby('title')
  ).then(data => sendJSON(res, data));
});

app.get('/content', (req, res) => {
  const graph = createGraphClient();
  getAll(graph, graph
    .api('/me/onenote/pages')
    .version('v1.0')
    .orderby('title')
  ).then(pages => getAllContent(graph, pages))
   .then(data => sendJSON(res, data))
});

app.listen(port, () => console.log(
  `OneNote Export server listening on port ${port}`));


const createGraphClient = () => MicrosoftGraph.Client.init({
  authProvider: done => done(null, accessToken),
  debugLogging: true,
});

const getAll = (graph: MicrosoftGraph.Client,
                req: MicrosoftGraph.GraphRequest): Promise<any[]> =>
  req.get().then(data => {
    const nextLink = data['@odata.nextLink'];
    const values: any[] = data.value;
    if (nextLink) {
      return getAll(graph, graph.api(nextLink))
        .then(nextValues => values.concat(nextValues));
    } else {
      return values;
    }
  });

const getAllContent = (graph: MicrosoftGraph.Client, 
                       pages: any[]): Promise<any[]> => {
  if (pages.length === 0) {
    return Promise.resolve([]);
  } else {
    const page = pages[0];
    return new Promise<Stream>((resolve, reject) => {
      graph.api(`/me/onenote/pages/${page.id}/content`)
        .getStream((err, stream) => stream ? resolve(stream) : reject(err))
    })
    .then(stream => new Promise<string>((resolve, reject) => {
      const chunks: any[] = [];
      stream.on("data", chunk => chunks.push(chunk));
      stream.on("end", () => resolve(Buffer.concat(chunks).toString()));
    }))
    .then(content => {
      page.content = content;
      return getAllContent(graph, pages.slice(1))
        .then(nextPages => [page].concat(nextPages));
    });
  }
};

const sendJSON = (res: Express.Response, data: object) => {
  res.set('Content-Type', 'application/json');
  res.send(JSON.stringify(data));
}