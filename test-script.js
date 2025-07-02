#!/usr/bin/env node

/**
 * Test script for the Azure Function msGraphConnector
 * 
 * This script performs browser-based authentication to get a token
 * and then tests the local Azure Function.
 * 
 * Prerequisites:
 * 1. Azure Function must be running locally (npm start)
 * 2. Azure AD app registration with appropriate permissions
 * 3. Environment variables set (see below)
 */

const http = require('http');
const url = require('url');
const { spawn } = require('child_process');
const axios = require('axios');

// Configuration - these should be set as environment variables
const config = {
    clientId: process.env.AZURE_CLIENT_ID || 'your-client-id-here',
    tenantId: process.env.AZURE_TENANT_ID || 'your-tenant-id-here',
    redirectUri: 'http://localhost:3000/auth/callback',
    scopes: [
        'https://graph.microsoft.com/Files.Read.All',
        'https://graph.microsoft.com/Sites.Read.All'
    ],
    functionUrl: process.env.FUNCTION_URL || 'http://localhost:7071/api/msGraphConnector'
};

console.log('🚀 Azure Function Test Script');
console.log('==============================');

// Validate configuration
if (config.clientId === 'your-client-id-here' || config.tenantId === 'your-tenant-id-here') {
    console.error('❌ Error: Please set AZURE_CLIENT_ID and AZURE_TENANT_ID environment variables');
    console.log('\nExample:');
    console.log('export AZURE_CLIENT_ID=12345678-1234-1234-1234-123456789abc');
    console.log('export AZURE_TENANT_ID=87654321-4321-4321-4321-cba987654321');
    console.log('npm run test-function');
    process.exit(1);
}

/**
 * Start a local HTTP server to handle the OAuth callback
 */
function startCallbackServer() {
    return new Promise((resolve, reject) => {
        const server = http.createServer((req, res) => {
            const parsedUrl = url.parse(req.url, true);
            
            if (parsedUrl.pathname === '/auth/callback') {
                const code = parsedUrl.query.code;
                const error = parsedUrl.query.error;
                
                if (error) {
                    res.writeHead(400, { 'Content-Type': 'text/html' });
                    res.end(`<h1>Authentication Error</h1><p>${error}: ${parsedUrl.query.error_description}</p>`);
                    reject(new Error(`Authentication error: ${error}`));
                    return;
                }
                
                if (code) {
                    res.writeHead(200, { 'Content-Type': 'text/html' });
                    res.end('<h1>Authentication Successful!</h1><p>You can close this window and return to the terminal.</p>');
                    server.close();
                    resolve(code);
                } else {
                    res.writeHead(400, { 'Content-Type': 'text/html' });
                    res.end('<h1>Authentication Error</h1><p>No authorization code received</p>');
                    reject(new Error('No authorization code received'));
                }
            } else {
                res.writeHead(404, { 'Content-Type': 'text/plain' });
                res.end('Not Found');
            }
        });
        
        server.listen(3000, (err) => {
            if (err) {
                reject(err);
            } else {
                console.log('📡 Callback server started on http://localhost:3000');
            }
        });
    });
}

/**
 * Open browser for authentication
 */
function openAuthUrl() {
    const authUrl = `https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0/authorize?` +
        `client_id=${config.clientId}&` +
        `response_type=code&` +
        `redirect_uri=${encodeURIComponent(config.redirectUri)}&` +
        `scope=${encodeURIComponent(config.scopes.join(' '))}&` +
        `response_mode=query&` +
        `prompt=select_account`;
    
    console.log('🌐 Opening browser for authentication...');
    console.log('📋 Auth URL:', authUrl);
    
    // Try to open browser
    const platform = process.platform;
    let command;
    
    if (platform === 'darwin') {
        command = 'open';
    } else if (platform === 'win32') {
        command = 'start';
    } else {
        command = 'xdg-open';
    }
    
    try {
        spawn(command, [authUrl], { detached: true, stdio: 'ignore' });
    } catch (error) {
        console.log('⚠️  Could not open browser automatically. Please open the URL above manually.');
    }
}

/**
 * Exchange authorization code for access token
 */
async function getAccessToken(authCode) {
    const tokenUrl = `https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0/token`;
    
    const params = new URLSearchParams({
        client_id: config.clientId,
        scope: config.scopes.join(' '),
        code: authCode,
        redirect_uri: config.redirectUri,
        grant_type: 'authorization_code'
    });
    
    try {
        console.log('🔑 Exchanging authorization code for access token...');
        const response = await axios.post(tokenUrl, params.toString(), {
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded'
            }
        });
        
        return response.data.access_token;
    } catch (error) {
        console.error('❌ Error getting access token:', error.response?.data || error.message);
        throw error;
    }
}

/**
 * Test the Azure Function
 */
async function testFunction(accessToken) {
    const searchTerm = process.argv[2] || 'test';
    
    console.log(`🔍 Testing function with search term: "${searchTerm}"`);
    console.log(`📡 Function URL: ${config.functionUrl}`);
    
    try {
        const response = await axios.post(config.functionUrl, 
            { searchTerm }, 
            {
                headers: {
                    'Authorization': `Bearer ${accessToken}`,
                    'Content-Type': 'application/json'
                },
                timeout: 30000 // 30 second timeout
            }
        );
        
        console.log('✅ Function Response:');
        console.log('Status:', response.status);
        console.log('Data:', JSON.stringify(response.data, null, 2));
        
    } catch (error) {
        console.error('❌ Function test failed:');
        if (error.response) {
            console.error('Status:', error.response.status);
            console.error('Data:', error.response.data);
        } else {
            console.error('Error:', error.message);
        }
        
        // Provide helpful debugging information
        if (error.code === 'ECONNREFUSED') {
            console.log('\n💡 Tip: Make sure the Azure Function is running locally with "npm start"');
        }
    }
}

/**
 * Main execution flow
 */
async function main() {
    try {
        // Start callback server
        const codePromise = startCallbackServer();
        
        // Open browser for authentication
        openAuthUrl();
        
        // Wait for authorization code
        console.log('⏳ Waiting for authentication...');
        const authCode = await codePromise;
        
        // Get access token
        const accessToken = await getAccessToken(authCode);
        console.log('✅ Access token obtained successfully');
        
        // Test the function
        await testFunction(accessToken);
        
        console.log('\n🎉 Test completed!');
        
    } catch (error) {
        console.error('❌ Test failed:', error.message);
        process.exit(1);
    }
}

// Check if we're running as main module
if (require.main === module) {
    main();
}

module.exports = { main, config };