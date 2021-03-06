/**
 * OfficeAuth 
 * @author Maksim Pahlberg
 * https://github.com/mpah
 * @created 18-03-2018
 *
 * OfficeAuth handles the authentification with the Office365 Backend through the oauth2 protocol
 * it requires a configured Azure AD application, defined in https://apps.dev.microsoft.com/ with 
 * oauth2AllowImplicitFlow": true
 *
 *
 */
class OfficeAuth {
	
	/**
	 * requests a token from office backend by navigating to url with parameters
	 */
	static requestToken(authServer, responseType, clientId, resource, replyUrl ) {
		//create the url
		var url = authServer + 
            "response_type=" + encodeURI(responseType) + "&" + 
            "client_id=" + encodeURI(clientId) + "&" + 
            "resource=" + encodeURI(resource) + "&" + 
            "redirect_uri=" + encodeURI(replyUrl); 
		//navigate to URL
		window.location = url; 
	}
	
	/**
	 * get the token from the url
	 */
	static getToken(){
		var accessCode = OfficeAuth.getJsonFromUrl().code;
		var params = { 
			grant_type: "authorization_code", 
			code: accessCode
		};  

		// Split the query string (after removing preceding '#'). 
		var queryStringParameters = OfficeAuth.splitQueryString(window.location.hash.substr(1)); 
		
		// Extract token from urlParameterExtraction object. Show token for debug purposes
		var token = queryStringParameters['access_token']; 
		return token;
	}
	
	/**
	 * parse json from url
	 */
	static getJsonFromUrl() {
	  var query = location.search.substr(1);
	  var result = {};
	  query.split("&").forEach(function(part) {
		var item = part.split("=");
		result[item[0]] = decodeURIComponent(item[1]);
	  });
	  return result;
	}
	
	/**
	 * split url
	 */
	static splitQueryString(queryStringFormattedString) { 
		var split = queryStringFormattedString.split('&'); 

		// If there are no parameters in URL, do nothing.
		if (split == "") {
		  return {};
		} 

		var results = {}; 

		// If there are parameters in URL, extract key/value pairs. 
		for (var i = 0; i < split.length; ++i) { 
		  var p = split[i].split('=', 2); 
		  if (p.length == 1) 
			results[p[0]] = ""; 
		  else 
			results[p[0]] = decodeURIComponent(p[1].replace(/\+/g, " ")); 
		} 

		return results; 
	}
	

}