const { app } = require('@azure/functions');
const { Client } = require('@microsoft/microsoft-graph-client');
const axios = require('axios');
const qs = require('querystring');
const { Buffer } = require('buffer');

// Initialize Microsoft Graph client
const initGraphClient = (accessToken) => {
    return Client.init({
        authProvider: (done) => {
            done(null, accessToken);
        }
    });
};

// Obtain OBO token using user's access token
const getOboToken = async (userAccessToken) => {
    const { TENANT_ID, CLIENT_ID, MICROSOFT_PROVIDER_AUTHENTICATION_SECRET } = process.env;
    const scope = 'https://graph.microsoft.com/.default';
    const oboTokenUrl = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
    const params = {
        client_id: CLIENT_ID,
        client_secret: MICROSOFT_PROVIDER_AUTHENTICATION_SECRET,
        grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
        assertion: userAccessToken,
        requested_token_use: 'on_behalf_of',
        scope: scope
    };
    try {
        const response = await axios.post(oboTokenUrl, qs.stringify(params), {
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
        });
        return response.data.access_token;
    } catch (error) {
        console.error('Error obtaining OBO token:', error.response?.data || error.message);
        throw error;
    }
};

// Fetch drive item content and convert to base64
const getDriveItemContent = async (client, driveId, itemId, name) => {
    try {
        const filePath = `/drives/${driveId}/items/${itemId}`;
        const downloadPath = filePath + `/content`;
        const fileStream = await client.api(downloadPath).getStream();
        let chunks = [];
        for await (let chunk of fileStream) {
            chunks.push(chunk);
        }
        const base64String = Buffer.concat(chunks).toString('base64');
        const file = await client.api(filePath).get();
        const mime_type = file.file.mimeType;
        const fileName = file.name;
        return { name: fileName, mime_type, content: base64String };
    } catch (error) {
        console.error('Error fetching drive content:', error);
        throw new Error(`Failed to fetch content for ${name}: ${error.message}`);
    }
};

app.http('msGraphConnector', {
    methods: ['GET', 'POST'],
    authLevel: 'function',
    handler: async (request, context) => {
        context.log(`Http function processed request for url "${request.url}"`);

        // Accept searchTerm from query or body
        let searchTerm = request.query.get('searchTerm');
        if (!searchTerm) {
            try {
                const body = await request.json();
                searchTerm = body.searchTerm;
            } catch {}
        }
        if (!searchTerm) {
            return { status: 400, body: 'Missing searchTerm parameter.' };
        }

        // Require Authorization header
        const authHeader = request.headers.get('authorization');
        if (!authHeader) {
            return { status: 400, body: 'Authorization header is missing' };
        }
        const bearerToken = authHeader.split(' ')[1];

        let accessToken;
        try {
            accessToken = await getOboToken(bearerToken);
        } catch (error) {
            return { status: 500, body: `Failed to obtain OBO token: ${error.message}` };
        }

        // Initialize Graph client
        const client = initGraphClient(accessToken);
        const requestBody = {
            requests: [
                {
                    entityTypes: ['driveItem'],
                    query: { queryString: searchTerm },
                    from: 0,
                    size: 10
                }
            ]
        };

        try {
            const list = await client.api('/search/query').post(requestBody);
            const processList = async () => {
                const results = [];
                await Promise.all(
                    (list.value[0].hitsContainers || []).map(async (container) => {
                        for (const hit of container.hits) {
                            if (hit.resource["@odata.type"] === "#microsoft.graph.driveItem") {
                                const { name, id } = hit.resource;
                                const driveId = hit.resource.parentReference.driveId;
                                const contents = await getDriveItemContent(client, driveId, id, name);
                                results.push(contents);
                            }
                        }
                    })
                );
                return results;
            };
            let results;
            if (!list.value[0].hitsContainers || list.value[0].hitsContainers[0].total === 0) {
                results = 'No results found';
            } else {
                results = await processList();
                results = { openaiFileResponse: results };
            }
            return { status: 200, jsonBody: results };
        } catch (error) {
            return { status: 500, body: `Error performing search or processing results: ${error.message}` };
        }
    }
});
