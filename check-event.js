var originalSaveButtonClickHandler = function(){};
jQuery(document).ready(function($) {
   // check existing event
   //This variable will hold a reference to the Announcements list items collection 
    var saveButton = $("[name$='diidIOSaveItem']") //gets form save button and ribbon save button
    if (saveButton.length > 0) {
      originalSaveButtonClickHandler = saveButton[0].onclick;  //save original function
    }
    $(saveButton).attr("onclick", "queryListItems()"); //change onclick to execute our custom validation function 
});


/**
 * Get the current editing form ID
 * @return current editing form ID
 */
function getParameterByName(name) {
  name = name.replace(/[\[]/, "\\[").replace(/[\]]/, "\\]");
  var regex = new RegExp("[\\?&]" + name + "=([^&#]*)"),
  results = regex.exec(location.search);
  return results === null ? "" : decodeURIComponent(results[1].replace(/\+/g, " "));
}

/**
 * Get time from a SharePoint Date / time field 
 * @field: title of input field
 * return @string of unix time stamp of input date field
 */
function getDateField(field){
  //Get date from input field
  var ddate = $(":input[title='" + field + "']").val();
  
  // date looks like DD/MM/YYYY
  var ddateArray = ddate.split('\/');
  var dmonth = ddateArray[0];
  var dday = ddateArray[1];
  var dyear= ddateArray[2];
  
  // Get Hour from input field
  var dDateID = $(":input[title='"+field+"']").attr("id");
  var dDateHours = $(":input[id='"+ dDateID + "Hours" +"']").val();

  // convert AM, PM hour from sharepoint input field to 24 hours
  var partOfDay = dDateHours.substr(dDateHours.length - 2);
  var hourPartOfDay = dDateHours.substring( 0, dDateHours.indexOf(partOfDay) );
  if (partOfDay == 'PM') {
    if (hourPartOfDay != 12) {
      hourPartOfDay = 12 + Number(hourPartOfDay);
    }
    else {
      hourPartOfDay = Number(hourPartOfDay);
    }
  }

  if (partOfDay == 'AM') {
    if (hourPartOfDay != 12) {
      hourPartOfDay = Number(hourPartOfDay);
    }
    else {
      hourPartOfDay = 0;
    }
  }
  
  // Get minute
  var dDateMinutes = $(":input[id='"+ dDateID + "Minutes" +"']").val();
  
  // convert to unix timestamp
  var departuredate = new Date(dyear, dmonth - 1, dday, hourPartOfDay, dDateMinutes,0).getTime()/1000;
  
  //return unix timestamp departuredate;
  return departuredate;
}

/**
 * get input field value
 * @field: title of input field
 * return string @data of input field
 */
function getInputField(field) {
  var data = $(":input[title='" + field + "']").val();
  return data;
}

/**
 * Get all the options from option field
 * @field: title of input field
 * return string @roomList
 */
function getAllElementInputForm(field) {
  var roomFieldID = $(":input[title='" + field + "']").attr("id");
  var roomOptions = document.getElementById(roomFieldID).options;
  var roomList = '';
  for(var z=0; z < roomOptions.length; z++){
    roomList = roomOptions[z].value + ', ' + roomList;
  }
  return roomList;
}

//This function fires when the query fails  
function onFailedCallback(sender, args) {  
  //Formulate HTML to display details of the error  
  var message = '<p>The request failed: <br>';  
      message += 'Message: ' + args.get_message() + '</p>';  
  //Display the details  
  document.getElementById("room_booked_message").innerHTML = message;
}  

//This function fires when the query completes successfully  
function onSucceededCallback(sender, args) {
  var saveRoom = true;
  var startTimeInput = getDateField("Start Date");
  var endTimeInput = getDateField("End Date");
  var conferenceRoomInput = getInputField("Conference Room");
  var roomAvailability = getAllElementInputForm("Conference Room");
  //Get an enumerator for the items in the list  
  var enumerator = returnedItems.getEnumerator();  

  //Loop through all the items  
  while (enumerator.moveNext()) {  
    
    var listItem = enumerator.get_current();
    
    // Get booked time
    var startTime = new Date(listItem.get_item('StartDate'));
    var endTime = new Date(listItem.get_item('_EndDate')); 
    
    //convert to unix timestamp for comparison
    var unixStartTime = startTime.getTime()/1000;
    var unixEndTime = endTime.getTime()/1000;
    
    // Get booked room conference     
    var roomConference =  listItem.get_item('Conference_x0020_Room');
        
    // Get user who booked the room
    var booker = listItem.get_item('Author').get_lookupValue();
    
    //Check time of new created booking room with booked rooms
    if ( (startTimeInput <= unixEndTime && startTimeInput >= unixStartTime)|| 
         (startTimeInput <= unixStartTime && endTimeInput >= unixEndTime) ||
         (endTimeInput <= unixEndTime && endTimeInput >= unixStartTime )) {
         roomAvailability = roomAvailability.replace(roomConference + ',','' );
         if (roomConference == conferenceRoomInput ) {
             //check if edit or create new booking
             if (listItem.get_item('ID') != getParameterByName('ID')) {// This is new form
                
                // There is an existing booking room time match with the new request booking
                saveRoom = false;
                
                // format date time
                var options = {  
                    weekday: "long", year: "numeric", month: "short",  
                    day: "numeric", hour: "2-digit", minute: "2-digit"  
                  }
                
                // Create a message  
                var roomBookedMessage = '<p><span class="ms-formvalidation">Room Booked</span>' + '<br/>'
                       + booker + '<br/>' +
                       'Start time: ' + startTime.toLocaleTimeString("en-us", options)  + '<br/>' +
                       'End time: ' + endTime.toLocaleTimeString("en-us", options) + '</p>'; 
              }   
         }
     }
  }
  
  if (saveRoom ) {
    // Save the new booking room
    originalSaveButtonClickHandler();
  }
  else {
    // There is an existing booking room time match with the request 
    // show the message to user
    var message = roomBookedMessage + '<p><span class="ms-formvalidation"> Available Rooms:</span> ' + roomAvailability + '</p>';
    document.getElementById("room_booked_message").innerHTML = message;
  }   

}  

/**
 * Get all books which have the EndDate greater than or equal to today
 * and the Reservation Status is equal to Avtive
 */
function queryListItems() { 
  //Get the current context  
  var context = new SP.ClientContext();  
  
  //Get the Announcements list. Alter this code to match the name of your list  
  var list = context.get_web().get_lists().getByTitle('Caltesting');  
  
  //Create a new CAML query    
  var caml = new SP.CamlQuery();  
  
  //Create the CAML that will return only items EndDate greater than or equal to today
  // and  Reservation_x0020_Status is equal to Avtive        
  var queryString = "<View><Query><Where><And>" +
                    "<Eq><FieldRef Name='Reservation_x0020_Status'/><Value Type='Text'>Active</Value></Eq>" + 
                    "<Geq><FieldRef Name='_EndDate'/><Value Type='DateTime'><Today/></Value></Geq>" + 
                    "</And></Where></Query></View>";        
 
 caml.set_viewXml(queryString);  
        
 //Specify the query and load the list oject  
 returnedItems = list.getItems(caml);  
 context.load(returnedItems);  
 
 //Run the query asynchronously, passing the functions to call when a response arrives  
 context.executeQueryAsync(onSucceededCallback, onFailedCallback);  
}  
