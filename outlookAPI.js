 /**
 * OutlookAPI 
 * @author Maksim Pahlberg
 * https://github.com/mpah
 * @created 18-03-2018
 *
 * OutlookAPI
 * uses the outlookAPI to request data from the Office backend, can be extended to handle all sorts of paths
 *
 */

class OutlookAPI {
	
	 /**
	 * get all calenders of authed user
	 */
	static getCalendars()
	{
		$.ajax({
		url: 'https://outlook.office.com/api/v2.0/me/calendars',
		type: 'GET',
		beforeSend: function (xhr) {
			xhr.setRequestHeader('Authorization', 'Bearer '+token);
		},
		data: {},
		success: function (response) 
		{ 
			//handle data here
			//OutlookAPI.createCalenderTable(response);
			OutlookAPI.createCalendarsTable(response);
		},
		error: function () { },
		});
	}
	
	/**
	 * get all events of a particular calendar defined by calenderId
	 */
	static getCalendar(calenderId)
	{
		$.ajax({
		url: 'https://outlook.office.com/api/v2.0/me/calendars/'+calenderId+'/events',
		type: 'GET',
		beforeSend: function (xhr) {
			xhr.setRequestHeader('Authorization', 'Bearer '+token);
		},
		data: {},
		success: function (response) 
		{ 
			//handle data here
			OutlookAPI.createCalenderTable(response);
		},
		error: function () { },
		});
	}
	
	
	/**
	 * helper to draw calendars list
	 */
	
	static createCalendarsTable(response)
	{
		document.writeln('<table style="width:100%">');
		response.value.forEach(function(entry) {
			document.writeln('<tr>');
			document.writeln('<td>'+entry.Name+'</td>');
			document.writeln('</tr>');
		});
		document.writeln('</table>');
	}
	
	/**
	 * helper to draw events list of a calendar
	 */
	static createCalenderTable(response)
	{
		document.writeln('<table style="width:100%">');
		response.value.forEach(function(entry) {
			document.writeln('<tr>');
			document.writeln('<td>'+entry.Start.DateTime+'</td>');
			document.writeln('<td>'+entry.End.DateTime+'</td>');
			document.writeln('<td>'+entry.Subject+'</td>');
			document.writeln('</tr>');
		});
		document.writeln('</table>');
	}
}