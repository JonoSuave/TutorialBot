/**
 * Helper function to call MS Graph API endpoint
 * using the authorization bearer token scheme
 */
export function callMSGraph(endpoint, access_token) {
	const headers = new Headers();
	const bearer = "Bearer " + access_token;
	headers.append("Authorization", bearer);
	const options = {
		method: "GET",
		headers: headers,
	};
	const graphEndpoint = endpoint;

	fetch(graphEndpoint, options).then(function (response) {
		//do something with response
	});
}
