# API scope defined in App Reg
export API_SCOPE="api://ff8d3a11-c9ef-490a-b95c-20cf34d23d4b"

# Login interactively (if you haven’t already)
# az login

# Grab an AAD access token for your Function’s scope
# This requires that AZ CLI has Authorized access to the App Registration
# (App Reg -> This App -> Expose an API -> Add a client application)
export TOKEN=$(az account get-access-token \
  --resource "$API_SCOPE" \
  --query accessToken -o tsv)

# only continue if success
if [ -n "$TOKEN" ]; then
  # print the token
  echo $TOKEN
  jwt $TOKEN

  # Call your local Function with that bearer token
  curl -H "Authorization: Bearer $TOKEN" \
    "http://localhost:7071/api/msGraphConnector?searchTerm=contoso"
fi

