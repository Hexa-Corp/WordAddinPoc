import express from 'express';
import axios from 'axios';
import cors from 'cors';

const app = express();

// Allow requests from localhost:3000 (your add-in)
app.use(cors({
  origin: 'https://localhost:3000'
}));

const tenantId = "2ef4acb6-d902-4515-8021-6eeeeb5d12bc";
const clientId = "be63a47a-1116-45dc-846b-659942fd924b";
const clientSecret = "ODP8Q~xk_3nYEZY5He1b4f6V3WezBpeaNUa9fbUp";

app.get('/getAppToken', async (req, res) => {
  try {
    const response = await axios.post(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
      new URLSearchParams({
        client_id: clientId,
        client_secret: clientSecret,
        scope: "https://graph.microsoft.com/.default",
        grant_type: "client_credentials"
      })
    );
    res.json({ token: response.data.access_token });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.listen(3001, () => console.log("Server running on port 3001"));



