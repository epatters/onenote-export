import Express from "express";

const app = Express();

app.get('/', (req, res) => res.send('OneNote Export'));

app.listen(3000, () => console.log(
  'OneNote Export server listening on port 3000'));